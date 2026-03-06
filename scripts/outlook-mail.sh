#!/bin/bash
# Outlook Mail Operations
# Usage: outlook-mail.sh [--account NAME] <command> [args]

BASE_DIR="$HOME/.outlook-mcp"

# Parse --account flag
ACCOUNT="${OUTLOOK_ACCOUNT:-default}"
if [ "$1" = "--account" ] || [ "$1" = "-a" ]; then
    ACCOUNT="$2"
    shift 2
fi

# Migrate legacy config to "default" subdirectory
if [ -f "$BASE_DIR/credentials.json" ] && [ ! -d "$BASE_DIR/default" ]; then
    mkdir -p "$BASE_DIR/default"
    mv "$BASE_DIR/config.json" "$BASE_DIR/default/" 2>/dev/null
    mv "$BASE_DIR/credentials.json" "$BASE_DIR/default/" 2>/dev/null
fi

# Validate account name to prevent directory traversal
if [[ ! "$ACCOUNT" =~ ^[a-zA-Z0-9_-]+$ ]]; then
    echo "Error: Invalid account name '$ACCOUNT'. Use only letters, numbers, hyphens, and underscores."
    exit 1
fi

CONFIG_DIR="$BASE_DIR/$ACCOUNT"
CREDS_FILE="$CONFIG_DIR/credentials.json"

# Validate count parameter is a positive integer
validate_count() {
    local val="$1"
    local default="$2"
    if [ -z "$val" ]; then
        echo "$default"
    elif [[ "$val" =~ ^[0-9]+$ ]]; then
        echo "$val"
    else
        echo "$default"
    fi
}

# Load token
ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE" 2>/dev/null)

if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
    echo "Error: Account '$ACCOUNT' not configured. Run: outlook-setup.sh --account $ACCOUNT"
    exit 1
fi

API="https://graph.microsoft.com/v1.0/me"

case "$1" in
    inbox)
        # List inbox messages
        COUNT=$(validate_count "${2:-}" 10)
        curl -s "$API/messages?\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], read: .value.isRead, id: .value.id[-20:]}'
        ;;
    
    unread)
        # List unread messages
        COUNT=$(validate_count "${2:-}" 20)
        curl -s "$API/messages?\$filter=isRead%20eq%20false&\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}'
        ;;
    
    search)
        # Search emails
        QUERY=$(jq -rn --arg q "$2" '$q | @uri')
        COUNT=$(validate_count "${3:-}" 20)
        curl -s "$API/messages?\$search=\"$QUERY\"&\$top=$COUNT&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}'
        ;;
    
    read)
        # Read specific email by ID (partial ID match - uses last 20 chars)
        MSG_ID="$2"
        # First find full ID (search by suffix)
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found. Use the ID shown in inbox/unread/search results."
            exit 1
        fi
        
        # Get message and extract text from HTML body
        curl -s "$API/messages/$FULL_ID?\$select=subject,from,receivedDateTime,body,toRecipients" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '{
                subject, 
                from: .from.emailAddress, 
                to: [.toRecipients[].emailAddress.address],
                date: .receivedDateTime,
                body: (if .body.contentType == "html" then (.body.content | gsub("<[^>]*>"; "") | gsub("\\s+"; " ") | gsub("&nbsp;"; " ") | .[0:2000]) else .body.content[0:2000] end)
            }'
        ;;
    
    mark-read)
        # Mark message as read
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"isRead": true}' | jq '{status: "marked as read", subject: .subject, id: .id[-20:]}'
        ;;
    
    folders)
        # List mail folders
        curl -s "$API/mailFolders" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value[] | {name: .displayName, total: .totalItemCount, unread: .unreadItemCount}'
        ;;
    
    stats)
        # Get inbox stats
        curl -s "$API/mailFolders/inbox" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '{folder: .displayName, total: .totalItemCount, unread: .unreadItemCount}'
        ;;
    
    send)
        # Send email: outlook-mail.sh send "to@email.com" "Subject" "Body"
        TO="$2"
        SUBJECT="$3"
        BODY="$4"
        
        if [ -z "$TO" ] || [ -z "$SUBJECT" ]; then
            echo "Usage: outlook-mail.sh send <to> <subject> <body>"
            exit 1
        fi
        
        JSON_PAYLOAD=$(jq -n --arg subj "$SUBJECT" --arg body "$BODY" --arg to "$TO" \
            '{message: {subject: $subj, body: {contentType: "Text", content: $body}, toRecipients: [{emailAddress: {address: $to}}]}}')

        RESULT=$(curl -s -w "\n%{http_code}" -X POST "$API/sendMail" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$JSON_PAYLOAD")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "202" ]; then
            jq -n --arg to "$TO" --arg subj "$SUBJECT" '{status: "sent", to: $to, subject: $subj}'
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    mark-unread)
        # Mark message as unread
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"isRead": false}' | jq '{status: "marked as unread", subject: .subject, id: .id[-20:]}'
        ;;
    
    delete)
        # Move message to trash
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s -X POST "$API/messages/$FULL_ID/move" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"destinationId": "deleteditems"}' | jq '{status: "moved to trash", subject: .subject, id: .id[-20:]}'
        ;;
    
    archive)
        # Move message to archive
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s -X POST "$API/messages/$FULL_ID/move" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"destinationId": "archive"}' | jq '{status: "archived", subject: .subject, id: .id[-20:]}'
        ;;
    
    flag)
        # Flag message as important
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"flag": {"flagStatus": "flagged"}}' | jq '{status: "flagged", subject: .subject, id: .id[-20:]}'
        ;;
    
    unflag)
        # Remove flag from message
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"flag": {"flagStatus": "notFlagged"}}' | jq '{status: "unflagged", subject: .subject, id: .id[-20:]}'
        ;;
    
    from)
        # List emails from specific sender (uses search - more reliable than filter)
        SENDER=$(jq -rn --arg s "$2" '$s | @uri')
        COUNT=$(validate_count "${3:-}" 20)
        curl -s "$API/messages?\$search=\"from:$SENDER\"&\$top=$COUNT&\$select=id,subject,from,receivedDateTime,isRead" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .value then (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], read: .value.isRead, id: .value.id[-20:]}) else {error: .error.message} end'
        ;;
    
    attachments)
        # List attachments for a message
        MSG_ID="$2"
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s "$API/messages/$FULL_ID/attachments" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value[] | {name: .name, size: .size, contentType: .contentType, id: .id}'
        ;;
    
    reply)
        # Reply to a message: outlook-mail.sh reply <id> "Reply body"
        MSG_ID="$2"
        BODY="$3"
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        RESULT=$(curl -s -w "\n%{http_code}" -X POST "$API/messages/$FULL_ID/reply" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$(jq -n --arg c "$BODY" '{comment: $c}')")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "202" ]; then
            jq -n --arg id "$MSG_ID" '{status: "reply sent", id: $id}'
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    move)
        # Move message to folder: outlook-mail.sh move <id> <folder>
        MSG_ID="$2"
        FOLDER="$3"
        
        if [ -z "$FOLDER" ]; then
            echo "Usage: outlook-mail.sh move <id> <folder>"
            echo "Use 'folders' command to see available folders"
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        # Get folder ID by name (case-insensitive)
        FOLDER_LOWER=$(echo "$FOLDER" | tr '[:upper:]' '[:lower:]')
        FOLDER_ID=$(curl -s "$API/mailFolders" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg fn "$FOLDER_LOWER" '.value[] | select((.displayName | ascii_downcase) == $fn) | .id' | head -1)
        
        if [ -z "$FOLDER_ID" ]; then
            echo "Folder not found: $FOLDER"
            echo "Available folders:"
            curl -s "$API/mailFolders" -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r '.value[].displayName'
            exit 1
        fi
        
        curl -s -X POST "$API/messages/$FULL_ID/move" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "{\"destinationId\": \"$FOLDER_ID\"}" | jq --arg f "$FOLDER" '{status: "moved", folder: $f, subject: .subject, id: .id[-20:]}'
        ;;
    
    draft)
        # Create a draft email (not sent)
        TO="$2"
        SUBJECT="$3"
        BODY="$4"
        
        if [ -z "$TO" ] || [ -z "$SUBJECT" ]; then
            echo "Usage: outlook-mail.sh draft <to> <subject> <body>"
            exit 1
        fi
        
        JSON_PAYLOAD=$(jq -n --arg subj "$SUBJECT" --arg body "$BODY" --arg to "$TO" \
            '{subject: $subj, body: {contentType: "Text", content: $body}, toRecipients: [{emailAddress: {address: $to}}]}')

        curl -s -X POST "$API/messages" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$JSON_PAYLOAD" | jq '{status: "draft created", subject: .subject, to: .toRecipients[0].emailAddress.address, id: .id[-20:]}'
        ;;
    
    drafts)
        # List draft emails
        COUNT=$(validate_count "${2:-}" 10)
        curl -s "$API/mailFolders/drafts/messages?\$top=$COUNT&\$select=id,subject,toRecipients,createdDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, to: (.value.toRecipients[0].emailAddress.address // "no recipient"), date: .value.createdDateTime[0:16], id: .value.id[-20:]}'
        ;;
    
    send-draft)
        # Send an existing draft
        MSG_ID="$2"
        
        # Search in drafts folder
        FULL_ID=$(curl -s "$API/mailFolders/drafts/messages?\$top=50&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Draft not found"
            exit 1
        fi
        
        RESULT=$(curl -s -w "\n%{http_code}" -X POST "$API/messages/$FULL_ID/send" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Length: 0")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "202" ]; then
            jq -n --arg id "$MSG_ID" '{status: "draft sent", id: $id}'
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    forward)
        # Forward an email: outlook-mail.sh forward <id> <to> [comment]
        MSG_ID="$2"
        TO="$3"
        COMMENT="${4:-}"
        
        if [ -z "$TO" ]; then
            echo "Usage: outlook-mail.sh forward <id> <to> [comment]"
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        JSON_PAYLOAD=$(jq -n --arg c "$COMMENT" --arg to "$TO" \
            '{comment: $c, toRecipients: [{emailAddress: {address: $to}}]}')

        RESULT=$(curl -s -w "\n%{http_code}" -X POST "$API/messages/$FULL_ID/forward" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$JSON_PAYLOAD")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "202" ]; then
            jq -n --arg to "$TO" --arg id "$MSG_ID" '{status: "forwarded", to: $to, id: $id}'
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    download)
        # Download an attachment: outlook-mail.sh download <msg-id> <attachment-name> [output-path]
        MSG_ID="$2"
        ATT_NAME="$3"
        OUTPUT="${4:-.}"

        # Validate output path is an existing directory
        if [ ! -d "$OUTPUT" ]; then
            echo "{\"error\": \"Output path is not an existing directory: $OUTPUT\"}"
            exit 1
        fi
        OUTPUT=$(cd "$OUTPUT" && pwd)

        if [ -z "$ATT_NAME" ]; then
            echo "Usage: outlook-mail.sh download <msg-id> <attachment-name> [output-path]"
            echo "Use 'attachments <id>' to see available attachments"
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        # Get attachment by name
        ATT_DATA=$(curl -s "$API/messages/$FULL_ID/attachments" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg aname "$ATT_NAME" '.value[] | select(.name == $aname)')
        
        if [ -z "$ATT_DATA" ]; then
            echo "Attachment not found: $ATT_NAME"
            echo "Available attachments:"
            curl -s "$API/messages/$FULL_ID/attachments" -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r '.value[].name'
            exit 1
        fi
        
        # Get content and decode
        ATT_ID=$(echo "$ATT_DATA" | jq -r '.id')
        CONTENT=$(curl -s "$API/messages/$FULL_ID/attachments/$ATT_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r '.contentBytes')
        
        # Sanitize filename: strip path traversal sequences and extract basename only
        SAFE_NAME=$(basename "$ATT_NAME" | sed 's/\.\.//g')
        if [ -z "$SAFE_NAME" ] || [ "$SAFE_NAME" = "." ] || [ "$SAFE_NAME" = ".." ]; then
            echo "{\"error\": \"Invalid attachment filename\"}"
            exit 1
        fi
        
        OUTPUT_FILE="$OUTPUT/$SAFE_NAME"
        echo "$CONTENT" | base64 -d > "$OUTPUT_FILE"
        
        if [ -f "$OUTPUT_FILE" ]; then
            SIZE=$(stat -c%s "$OUTPUT_FILE" 2>/dev/null || stat -f%z "$OUTPUT_FILE")
            echo "{\"status\": \"downloaded\", \"file\": \"$OUTPUT_FILE\", \"size\": $SIZE}"
        else
            echo "{\"error\": \"Failed to save file\"}"
            exit 1
        fi
        ;;
    
    create-folder)
        # Create a new mail folder
        FOLDER_NAME="$2"
        PARENT="${3:-}"
        
        if [ -z "$FOLDER_NAME" ]; then
            echo "Usage: outlook-mail.sh create-folder <name> [parent-folder]"
            exit 1
        fi
        
        if [ -n "$PARENT" ]; then
            # Get parent folder ID
            PARENT_LOWER=$(echo "$PARENT" | tr '[:upper:]' '[:lower:]')
            PARENT_ID=$(curl -s "$API/mailFolders" \
                -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg fn "$PARENT_LOWER" '.value[] | select((.displayName | ascii_downcase) == $fn) | .id' | head -1)
            
            if [ -z "$PARENT_ID" ]; then
                echo "Parent folder not found: $PARENT"
                exit 1
            fi
            
            curl -s -X POST "$API/mailFolders/$PARENT_ID/childFolders" \
                -H "Authorization: Bearer $ACCESS_TOKEN" \
                -H "Content-Type: application/json" \
                -d "$(jq -n --arg n "$FOLDER_NAME" '{displayName: $n}')" | jq --arg p "$PARENT" '{status: "folder created", name: .displayName, parent: $p, id: .id[-20:]}'
        else
            curl -s -X POST "$API/mailFolders" \
                -H "Authorization: Bearer $ACCESS_TOKEN" \
                -H "Content-Type: application/json" \
                -d "$(jq -n --arg n "$FOLDER_NAME" '{displayName: $n}')" | jq '{status: "folder created", name: .displayName, id: .id[-20:]}'
        fi
        ;;
    
    delete-folder)
        # Delete a mail folder
        FOLDER_NAME="$2"
        
        if [ -z "$FOLDER_NAME" ]; then
            echo "Usage: outlook-mail.sh delete-folder <name>"
            exit 1
        fi
        
        FOLDER_LOWER=$(echo "$FOLDER_NAME" | tr '[:upper:]' '[:lower:]')
        FOLDER_ID=$(curl -s "$API/mailFolders" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg fn "$FOLDER_LOWER" '.value[] | select((.displayName | ascii_downcase) == $fn) | .id' | head -1)
        
        if [ -z "$FOLDER_ID" ]; then
            echo "Folder not found: $FOLDER_NAME"
            exit 1
        fi
        
        RESULT=$(curl -s -w "\n%{http_code}" -X DELETE "$API/mailFolders/$FOLDER_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "204" ]; then
            jq -n --arg n "$FOLDER_NAME" '{status: "folder deleted", name: $n}'
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    bulk-read)
        # Mark multiple messages as read: outlook-mail.sh bulk-read <id1> <id2> ...
        shift
        if [ $# -eq 0 ]; then
            echo "Usage: outlook-mail.sh bulk-read <id1> <id2> <id3> ..."
            exit 1
        fi
        
        SUCCESS=0
        FAILED=0
        
        for MSG_ID in "$@"; do
            FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
                -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
            
            if [ -n "$FULL_ID" ]; then
                curl -s -X PATCH "$API/messages/$FULL_ID" \
                    -H "Authorization: Bearer $ACCESS_TOKEN" \
                    -H "Content-Type: application/json" \
                    -d '{"isRead": true}' > /dev/null
                SUCCESS=$((SUCCESS + 1))
            else
                FAILED=$((FAILED + 1))
            fi
        done
        
        echo "{\"status\": \"bulk operation complete\", \"marked_read\": $SUCCESS, \"not_found\": $FAILED}"
        ;;
    
    bulk-delete)
        # Delete multiple messages: outlook-mail.sh bulk-delete <id1> <id2> ...
        shift
        if [ $# -eq 0 ]; then
            echo "Usage: outlook-mail.sh bulk-delete <id1> <id2> <id3> ..."
            exit 1
        fi
        
        SUCCESS=0
        FAILED=0
        
        for MSG_ID in "$@"; do
            FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
                -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
            
            if [ -n "$FULL_ID" ]; then
                curl -s -X POST "$API/messages/$FULL_ID/move" \
                    -H "Authorization: Bearer $ACCESS_TOKEN" \
                    -H "Content-Type: application/json" \
                    -d '{"destinationId": "deleteditems"}' > /dev/null
                SUCCESS=$((SUCCESS + 1))
            else
                FAILED=$((FAILED + 1))
            fi
        done
        
        echo "{\"status\": \"bulk delete complete\", \"deleted\": $SUCCESS, \"not_found\": $FAILED}"
        ;;
    
    categories)
        # List available categories (like Gmail labels)
        curl -s "https://graph.microsoft.com/v1.0/me/outlook/masterCategories" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value[] | {name: .displayName, color: .color, id: .id[0:8]}'
        ;;
    
    categorize)
        # Add category to message: outlook-mail.sh categorize <id> <category-name>
        MSG_ID="$2"
        CATEGORY="$3"
        
        if [ -z "$CATEGORY" ]; then
            echo "Usage: outlook-mail.sh categorize <id> <category-name>"
            echo "Use 'categories' to see available categories"
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id,categories" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        # Get current categories and add new one
        CURRENT_CATS=$(curl -s "$API/messages/$FULL_ID?\$select=categories" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.categories')

        NEW_CATS=$(echo "$CURRENT_CATS" | jq --arg cat "$CATEGORY" '. + [$cat]')

        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$(jq -n --argjson cats "$NEW_CATS" '{categories: $cats}')" | jq '{status: "categorized", subject: .subject, categories: .categories, id: .id[-20:]}'
        ;;
    
    uncategorize)
        # Remove all categories from message
        MSG_ID="$2"
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        curl -s -X PATCH "$API/messages/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d '{"categories": []}' | jq '{status: "categories removed", subject: .subject, id: .id[-20:]}'
        ;;
    
    rules)
        # Show current auto-categorize rules
        RULES_FILE="$CONFIG_DIR/rules.json"
        if [ ! -f "$RULES_FILE" ]; then
            echo "{\"info\": \"No rules configured. Use add-rule to create rules.\"}"
            exit 0
        fi
        jq '.' "$RULES_FILE"
        ;;

    add-rule)
        # Add an auto-categorize rule: outlook-mail.sh add-rule <field> <pattern> <category>
        MATCH_FIELD="$2"
        PATTERN="$3"
        CATEGORY="$4"

        if [ -z "$MATCH_FIELD" ] || [ -z "$PATTERN" ] || [ -z "$CATEGORY" ]; then
            echo "Usage: outlook-mail.sh add-rule <field> <pattern> <category>"
            echo "Fields: from, subject"
            echo "Example: outlook-mail.sh add-rule from @github.com Dev"
            echo "Example: outlook-mail.sh add-rule subject invoice Finance"
            exit 1
        fi

        case "$MATCH_FIELD" in
            from|subject) ;;
            *)
                echo "Error: field must be 'from' or 'subject'"
                exit 1
                ;;
        esac

        RULES_FILE="$CONFIG_DIR/rules.json"
        if [ ! -f "$RULES_FILE" ]; then
            echo '{"rules":[]}' > "$RULES_FILE"
            chmod 600 "$RULES_FILE"
        fi

        jq --arg f "$MATCH_FIELD" --arg p "$PATTERN" --arg c "$CATEGORY" \
            '.rules += [{match: $f, pattern: $p, category: $c}]' "$RULES_FILE" > "$RULES_FILE.tmp" \
            && mv "$RULES_FILE.tmp" "$RULES_FILE"
        jq -n --arg f "$MATCH_FIELD" --arg p "$PATTERN" --arg c "$CATEGORY" \
            '{status: "rule added", match: $f, pattern: $p, category: $c}'
        ;;

    remove-rule)
        # Remove a rule by index (0-based): outlook-mail.sh remove-rule <index>
        RULE_INDEX="$2"

        if [ -z "$RULE_INDEX" ]; then
            echo "Usage: outlook-mail.sh remove-rule <index>"
            echo "Use 'rules' to see current rules with their indices"
            exit 1
        fi

        RULES_FILE="$CONFIG_DIR/rules.json"
        if [ ! -f "$RULES_FILE" ]; then
            echo "{\"error\": \"No rules file found\"}"
            exit 1
        fi

        RULE_COUNT=$(jq '.rules | length' "$RULES_FILE")
        if [ "$RULE_INDEX" -ge "$RULE_COUNT" ] 2>/dev/null; then
            echo "{\"error\": \"Rule index out of range. Max index: $((RULE_COUNT - 1))\"}"
            exit 1
        fi

        REMOVED=$(jq --argjson i "$RULE_INDEX" '.rules[$i]' "$RULES_FILE")
        jq --argjson i "$RULE_INDEX" 'del(.rules[$i])' "$RULES_FILE" > "$RULES_FILE.tmp" \
            && mv "$RULES_FILE.tmp" "$RULES_FILE"
        echo "$REMOVED" | jq '{status: "rule removed"} + .'
        ;;

    auto-categorize)
        # Apply rules to recent uncategorized emails: outlook-mail.sh auto-categorize [count]
        COUNT=$(validate_count "${2:-}" 50)
        RULES_FILE="$CONFIG_DIR/rules.json"

        if [ ! -f "$RULES_FILE" ]; then
            echo "{\"error\": \"No rules configured. Use add-rule to create rules.\"}"
            exit 1
        fi

        RULE_COUNT=$(jq '.rules | length' "$RULES_FILE")
        if [ "$RULE_COUNT" -eq 0 ]; then
            echo "{\"error\": \"No rules defined. Use add-rule to create rules.\"}"
            exit 1
        fi

        # Fetch recent emails with their current categories
        MESSAGES=$(curl -s "$API/messages?\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,categories" \
            -H "Authorization: Bearer $ACCESS_TOKEN")

        MSG_COUNT=$(echo "$MESSAGES" | jq '.value | length')
        CATEGORIZED=0
        SKIPPED=0
        ALREADY=0

        for i in $(seq 0 $((MSG_COUNT - 1))); do
            MSG=$(echo "$MESSAGES" | jq --argjson i "$i" '.value[$i]')
            MSG_ID=$(echo "$MSG" | jq -r '.id')
            MSG_SUBJECT=$(echo "$MSG" | jq -r '.subject // ""' | tr '[:upper:]' '[:lower:]')
            MSG_FROM=$(echo "$MSG" | jq -r '.from.emailAddress.address // ""' | tr '[:upper:]' '[:lower:]')
            EXISTING_CATS=$(echo "$MSG" | jq '.categories')

            # Collect matching categories from all rules
            NEW_CATS="$EXISTING_CATS"
            MATCHED=false

            for r in $(seq 0 $((RULE_COUNT - 1))); do
                RULE=$(jq --argjson r "$r" '.rules[$r]' "$RULES_FILE")
                MATCH_FIELD=$(echo "$RULE" | jq -r '.match')
                PATTERN=$(echo "$RULE" | jq -r '.pattern' | tr '[:upper:]' '[:lower:]')
                CATEGORY=$(echo "$RULE" | jq -r '.category')

                # Check if category already applied
                HAS_CAT=$(echo "$NEW_CATS" | jq --arg c "$CATEGORY" 'map(select(. == $c)) | length')
                if [ "$HAS_CAT" -gt 0 ]; then
                    continue
                fi

                # Match against the appropriate field
                MATCH_VALUE=""
                case "$MATCH_FIELD" in
                    from) MATCH_VALUE="$MSG_FROM" ;;
                    subject) MATCH_VALUE="$MSG_SUBJECT" ;;
                esac

                if echo "$MATCH_VALUE" | grep -qi "$PATTERN" 2>/dev/null; then
                    NEW_CATS=$(echo "$NEW_CATS" | jq --arg c "$CATEGORY" '. + [$c]')
                    MATCHED=true
                fi
            done

            if [ "$MATCHED" = true ]; then
                # Apply the updated categories
                curl -s -X PATCH "$API/messages/$MSG_ID" \
                    -H "Authorization: Bearer $ACCESS_TOKEN" \
                    -H "Content-Type: application/json" \
                    -d "$(jq -n --argjson cats "$NEW_CATS" '{categories: $cats}')" > /dev/null
                CATEGORIZED=$((CATEGORIZED + 1))
                MSG_SUBJ_DISPLAY=$(echo "$MSG" | jq -r '.subject // "(no subject)"')
                APPLIED_CATS=$(echo "$NEW_CATS" | jq -c '.')
                jq -n --arg s "$MSG_SUBJ_DISPLAY" --argjson c "$NEW_CATS" '{applied: $s, categories: $c}'
            else
                # Check if it already had categories
                EXISTING_LEN=$(echo "$EXISTING_CATS" | jq 'length')
                if [ "$EXISTING_LEN" -gt 0 ]; then
                    ALREADY=$((ALREADY + 1))
                else
                    SKIPPED=$((SKIPPED + 1))
                fi
            fi
        done

        jq -n --argjson c "$CATEGORIZED" --argjson s "$SKIPPED" --argjson a "$ALREADY" --argjson t "$MSG_COUNT" \
            '{status: "auto-categorize complete", scanned: $t, categorized: $c, no_match: $s, already_categorized: $a}'
        ;;

    focused)
        # List focused inbox (important emails)
        COUNT=$(validate_count "${2:-}" 10)
        curl -s "$API/messages?\$filter=inferenceClassification%20eq%20'focused'&\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .value then (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}) else {info: "Focused inbox not available or empty"} end'
        ;;
    
    other)
        # List "other" inbox (non-focused emails)
        COUNT=$(validate_count "${2:-}" 10)
        curl -s "$API/messages?\$filter=inferenceClassification%20eq%20'other'&\$top=$COUNT&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq 'if .value then (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}) else {info: "Other inbox not available or empty"} end'
        ;;
    
    thread)
        # List emails in same conversation/thread (by subject keyword)
        MSG_ID="$2"
        
        FULL_ID=$(curl -s "$API/messages?\$top=100&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg mid "$MSG_ID" '.value[] | select(.id | endswith($mid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Message not found"
            exit 1
        fi
        
        # Get subject, clean prefixes
        SUBJECT=$(curl -s "$API/messages/$FULL_ID?\$select=subject" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r '.subject' | sed 's/^RE: //i' | sed 's/^FW: //i' | sed 's/^Fwd: //i')
        
        # Get longest word as keyword (more specific)
        KEYWORD=$(echo "$SUBJECT" | tr ' ' '\n' | awk '{print length, $0}' | sort -rn | head -1 | cut -d' ' -f2)
        
        if [ -z "$KEYWORD" ] || [ ${#KEYWORD} -lt 4 ]; then
            KEYWORD=$(echo "$SUBJECT" | cut -d' ' -f1)
        fi
        
        echo "Searching thread by keyword: $KEYWORD"

        # Search by single keyword
        KEYWORD_ENC=$(jq -rn --arg k "$KEYWORD" '$k | @uri')
        curl -s "$API/messages?\$search=\"$KEYWORD_ENC\"&\$top=20&\$select=id,subject,from,receivedDateTime" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, from: .value.from.emailAddress.address, date: .value.receivedDateTime[0:16], id: .value.id[-20:]}'
        ;;
    
    *)
        echo "Usage: outlook-mail.sh <command> [args]"
        echo ""
        echo "READING:"
        echo "  inbox [count]             - List latest emails (default: 10)"
        echo "  unread [count]            - List unread emails"
        echo "  focused [count]           - Focused/important inbox"
        echo "  other [count]             - Other/low-priority inbox"
        echo "  search \"query\" [count]   - Search emails"
        echo "  from <email> [count]      - List emails from sender"
        echo "  read <id>                 - Read email content"
        echo "  thread <id>               - View conversation thread"
        echo "  attachments <id>          - List email attachments"
        echo ""
        echo "MANAGING:"
        echo "  mark-read <id>            - Mark as read"
        echo "  mark-unread <id>          - Mark as unread"
        echo "  flag <id>                 - Flag as important"
        echo "  unflag <id>               - Remove flag"
        echo "  delete <id>               - Move to trash"
        echo "  archive <id>              - Move to archive"
        echo "  move <id> <folder>        - Move to folder"
        echo ""
        echo "CATEGORIES (like Gmail labels):"
        echo "  categories                - List available categories"
        echo "  categorize <id> <name>    - Add category to email"
        echo "  uncategorize <id>         - Remove categories"
        echo ""
        echo "AUTO-CATEGORIZE:"
        echo "  rules                     - Show current rules"
        echo "  add-rule <field> <pattern> <category>"
        echo "                            - Add rule (field: from, subject)"
        echo "  remove-rule <index>       - Remove rule by index"
        echo "  auto-categorize [count]   - Apply rules to recent emails"
        echo ""
        echo "SENDING:"
        echo "  send <to> <subj> <body>   - Send new email"
        echo "  reply <id> \"body\"         - Reply to email"
        echo "  forward <id> <to> [msg]   - Forward email"
        echo ""
        echo "DRAFTS:"
        echo "  draft <to> <subj> <body>  - Create draft (not sent)"
        echo "  drafts [count]            - List drafts"
        echo "  send-draft <id>           - Send a draft"
        echo ""
        echo "ATTACHMENTS:"
        echo "  attachments <id>          - List attachments"
        echo "  download <id> <name> [path] - Download attachment"
        echo ""
        echo "FOLDERS:"
        echo "  folders                   - List mail folders"
        echo "  create-folder <name> [parent] - Create folder"
        echo "  delete-folder <name>      - Delete folder"
        echo ""
        echo "BULK OPERATIONS:"
        echo "  bulk-read <id1> <id2>...  - Mark multiple as read"
        echo "  bulk-delete <id1> <id2>...  - Delete multiple"
        echo ""
        echo "INFO:"
        echo "  stats                     - Inbox statistics"
        ;;
esac
