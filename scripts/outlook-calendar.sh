#!/bin/bash
# Outlook Calendar Operations
# Usage: outlook-calendar.sh [--account NAME] <command> [args]

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

# Detect timezone: use OUTLOOK_TZ env var, or system timezone, or fallback to UTC
if [ -n "$OUTLOOK_TZ" ]; then
    TIMEZONE="$OUTLOOK_TZ"
elif [ -f /etc/timezone ]; then
    TIMEZONE=$(cat /etc/timezone)
elif command -v timedatectl &> /dev/null; then
    TIMEZONE=$(timedatectl show --property=Timezone --value 2>/dev/null)
elif [ -L /etc/localtime ]; then
    TIMEZONE=$(readlink /etc/localtime | sed 's|.*/zoneinfo/||')
else
    # macOS fallback
    TIMEZONE=$(ls -l /etc/localtime 2>/dev/null | sed 's|.*/zoneinfo/||')
fi
# Validate timezone to prevent HTTP header injection
TIMEZONE="${TIMEZONE:-UTC}"
if [[ ! "$TIMEZONE" =~ ^[a-zA-Z0-9/_+-]+$ ]]; then
    echo "Error: Invalid timezone '$TIMEZONE'. Use standard IANA format (e.g., America/New_York)."
    exit 1
fi

case "$1" in
    events)
        # List upcoming events
        COUNT=$(validate_count "${2:-}" 10)
        curl -s "$API/calendar/events?\$top=$COUNT&\$orderby=start/dateTime%20desc&\$select=id,subject,start,end,location,isAllDay" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq '.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, start: .value.start.dateTime[0:16], end: .value.end.dateTime[0:16], location: (.value.location.displayName // ""), id: .value.id[-20:]}'
        ;;
    
    today)
        # List today's events
        TODAY_START=$(date -u +"%Y-%m-%dT00:00:00Z")
        TODAY_END=$(date -u +"%Y-%m-%dT23:59:59Z")
        curl -s "$API/calendarView?startDateTime=$TODAY_START&endDateTime=$TODAY_END&\$orderby=start/dateTime&\$select=id,subject,start,end,location" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq 'if .value then (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, start: .value.start.dateTime[0:16], end: .value.end.dateTime[0:16], location: (.value.location.displayName // ""), id: .value.id[-20:]}) else {error: .error.message} end'
        ;;
    
    week)
        # List this week's events
        WEEK_START=$(date -u +"%Y-%m-%dT00:00:00Z")
        WEEK_END=$(date -u -d "+7 days" +"%Y-%m-%dT23:59:59Z" 2>/dev/null || date -u -v+7d +"%Y-%m-%dT23:59:59Z")
        curl -s "$API/calendarView?startDateTime=$WEEK_START&endDateTime=$WEEK_END&\$orderby=start/dateTime&\$select=id,subject,start,end,location,isAllDay" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq 'if .value then (.value | to_entries | .[] | {n: (.key + 1), subject: .value.subject, start: .value.start.dateTime[0:16], end: .value.end.dateTime[0:16], location: (.value.location.displayName // ""), id: .value.id[-20:]}) else {error: .error.message} end'
        ;;
    
    read)
        # Read event details
        EVENT_ID="$2"
        FULL_ID=$(curl -s "$API/calendar/events?\$top=50&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg eid "$EVENT_ID" '.value[] | select(.id | endswith($eid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Event not found"
            exit 1
        fi
        
        curl -s "$API/calendar/events/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$TIMEZONE\"" | jq '{
                subject,
                start: .start.dateTime,
                end: .end.dateTime,
                location: .location.displayName,
                body: (if .body.contentType == "html" then (.body.content | gsub("<[^>]*>"; "") | gsub("\\s+"; " ")[0:500]) else .body.content[0:500] end),
                attendees: [.attendees[]?.emailAddress.address],
                isOnline: .isOnlineMeeting,
                link: .onlineMeeting.joinUrl
            }'
        ;;
    
    create)
        # Create event: outlook-calendar.sh create "Subject" "2026-01-26T10:00" "2026-01-26T11:00" [location]
        SUBJECT="$2"
        START="$3"
        END="$4"
        LOCATION="${5:-}"
        
        if [ -z "$SUBJECT" ] || [ -z "$START" ] || [ -z "$END" ]; then
            echo "Usage: outlook-calendar.sh create <subject> <start> <end> [location]"
            echo "Date format: YYYY-MM-DDTHH:MM (e.g., 2026-01-26T10:00)"
            exit 1
        fi
        
        if [ -n "$LOCATION" ]; then
            JSON_PAYLOAD=$(jq -n --arg subj "$SUBJECT" --arg s "$START" --arg e "$END" --arg tz "$TIMEZONE" --arg loc "$LOCATION" \
                '{subject: $subj, start: {dateTime: $s, timeZone: $tz}, end: {dateTime: $e, timeZone: $tz}, location: {displayName: $loc}}')
        else
            JSON_PAYLOAD=$(jq -n --arg subj "$SUBJECT" --arg s "$START" --arg e "$END" --arg tz "$TIMEZONE" \
                '{subject: $subj, start: {dateTime: $s, timeZone: $tz}, end: {dateTime: $e, timeZone: $tz}}')
        fi

        curl -s -X POST "$API/calendar/events" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$JSON_PAYLOAD" | jq '{status: "event created", subject: .subject, start: .start.dateTime[0:16], end: .end.dateTime[0:16], id: .id[-20:]}'
        ;;
    
    quick)
        # Quick event (1 hour from now or specified time)
        SUBJECT="$2"
        START_TIME="${3:-}"
        
        if [ -z "$SUBJECT" ]; then
            echo "Usage: outlook-calendar.sh quick <subject> [start-time]"
            echo "If no time given, creates 1-hour event starting now"
            exit 1
        fi
        
        if [ -z "$START_TIME" ]; then
            START=$(date +"%Y-%m-%dT%H:%M")
            END=$(date -d "+1 hour" +"%Y-%m-%dT%H:%M" 2>/dev/null || date -v+1H +"%Y-%m-%dT%H:%M")
        else
            START="$START_TIME"
            # Parse and add 1 hour
            END=$(date -d "$START_TIME + 1 hour" +"%Y-%m-%dT%H:%M" 2>/dev/null || echo "$START_TIME")
        fi
        
        JSON_PAYLOAD=$(jq -n --arg subj "$SUBJECT" --arg s "$START" --arg e "$END" --arg tz "$TIMEZONE" \
            '{subject: $subj, start: {dateTime: $s, timeZone: $tz}, end: {dateTime: $e, timeZone: $tz}}')

        curl -s -X POST "$API/calendar/events" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$JSON_PAYLOAD" | jq '{status: "quick event created", subject: .subject, start: .start.dateTime[0:16], end: .end.dateTime[0:16], id: .id[-20:]}'
        ;;
    
    delete)
        # Delete event
        EVENT_ID="$2"
        FULL_ID=$(curl -s "$API/calendar/events?\$top=50&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg eid "$EVENT_ID" '.value[] | select(.id | endswith($eid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Event not found"
            exit 1
        fi
        
        RESULT=$(curl -s -w "\n%{http_code}" -X DELETE "$API/calendar/events/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN")
        
        HTTP_CODE=$(echo "$RESULT" | tail -1)
        if [ "$HTTP_CODE" = "204" ]; then
            echo "{\"status\": \"event deleted\", \"id\": \"$EVENT_ID\"}"
        else
            echo "$RESULT" | head -n -1 | jq '.error // .'
        fi
        ;;
    
    update)
        # Update event: outlook-calendar.sh update <id> <field> <value>
        EVENT_ID="$2"
        FIELD="$3"
        VALUE="$4"
        
        if [ -z "$FIELD" ] || [ -z "$VALUE" ]; then
            echo "Usage: outlook-calendar.sh update <id> <field> <value>"
            echo "Fields: subject, location, start, end"
            exit 1
        fi
        
        FULL_ID=$(curl -s "$API/calendar/events?\$top=50&\$select=id" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq -r --arg eid "$EVENT_ID" '.value[] | select(.id | endswith($eid)) | .id' | head -1)
        
        if [ -z "$FULL_ID" ]; then
            echo "Event not found"
            exit 1
        fi
        
        case "$FIELD" in
            subject)
                PATCH_BODY=$(jq -n --arg v "$VALUE" '{subject: $v}')
                ;;
            location)
                PATCH_BODY=$(jq -n --arg v "$VALUE" '{location: {displayName: $v}}')
                ;;
            start)
                PATCH_BODY=$(jq -n --arg v "$VALUE" --arg tz "$TIMEZONE" '{start: {dateTime: $v, timeZone: $tz}}')
                ;;
            end)
                PATCH_BODY=$(jq -n --arg v "$VALUE" --arg tz "$TIMEZONE" '{end: {dateTime: $v, timeZone: $tz}}')
                ;;
            *)
                echo "Unknown field: $FIELD"
                exit 1
                ;;
        esac
        
        curl -s -X PATCH "$API/calendar/events/$FULL_ID" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$PATCH_BODY" | jq '{status: "event updated", subject: .subject, start: .start.dateTime[0:16], end: .end.dateTime[0:16], id: .id[-20:]}'
        ;;
    
    calendars)
        # List all calendars
        curl -s "$API/calendars" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq '.value[] | {name: .name, color: .color, canEdit: .canEdit, id: .id[-20:]}'
        ;;
    
    free)
        # Check free/busy for a time range
        START="$2"
        END="$3"
        
        if [ -z "$START" ] || [ -z "$END" ]; then
            echo "Usage: outlook-calendar.sh free <start> <end>"
            echo "Date format: YYYY-MM-DDTHH:MM"
            exit 1
        fi
        
        curl -s "$API/calendarView?startDateTime=${START}:00Z&endDateTime=${END}:00Z&\$select=subject,start,end" \
            -H "Authorization: Bearer $ACCESS_TOKEN" | jq --arg s "$START" --arg e "$END" 'if (.value | length) == 0 then {status: "free", start: $s, end: $e} else {status: "busy", events: [.value[].subject]} end'
        ;;
    
    *)
        echo "Usage: outlook-calendar.sh <command> [args]"
        echo ""
        echo "VIEW:"
        echo "  events [count]            - List upcoming events"
        echo "  today                     - Today's events"
        echo "  week                      - This week's events"
        echo "  read <id>                 - Event details"
        echo "  calendars                 - List all calendars"
        echo "  free <start> <end>        - Check availability"
        echo ""
        echo "CREATE:"
        echo "  create <subj> <start> <end> [loc] - Create event"
        echo "  quick <subject> [time]    - Quick 1-hour event"
        echo ""
        echo "MANAGE:"
        echo "  update <id> <field> <val> - Update event"
        echo "  delete <id>               - Delete event"
        echo ""
        echo "Date format: YYYY-MM-DDTHH:MM (e.g., 2026-01-26T10:00)"
        ;;
esac
