import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from collections import defaultdict
import sys
import re
import os
import csv
from urllib.parse import urljoin
from bs4 import BeautifulSoup
import re


def parse_date_input(date_str):
    """Parse date input in m/d/yy format."""
    try:
        # Parse as naive datetime, then convert to match timezone-aware datetimes
        dt = datetime.strptime(date_str.strip(), "%m/%d/%y")
        # Set to UTC for comparison
        return dt.replace(tzinfo=timezone.utc)
    except ValueError:
        print(f"Invalid date format: {date_str}. Please use m/d/yy format.")
        sys.exit(1)


def parse_iso_datetime(iso_string):
    """Parse ISO 8601 datetime string."""
    # Handle both with and without timezone
    if iso_string.endswith('Z'):
        iso_string = iso_string[:-1] + '+00:00'
    try:
        return datetime.fromisoformat(iso_string)
    except ValueError:
        return None


def fetch_rss_feed(url):
    """Fetch RSS feed from URL."""
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        return response.text
    except requests.RequestException as e:
        print(f"Error fetching RSS feed: {e}")
        sys.exit(1)


def fetch_event_detail_page(url):
    """Fetch event detail page and extract speaker and location info."""
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        speaker = None
        location = None
        youtube_link = None
        
        # Extract speaker name from .pageTitle > .subtitle
        page_title = soup.find('div', class_='pageTitle')
        if page_title:
            subtitle = page_title.find('div', class_='subtitle')
            if subtitle:
                speaker = subtitle.get_text(strip=True)
        
        # Extract location from .event-detail-float .place
        event_detail_float = soup.find('div', class_='event-detail-float')
        if event_detail_float:
            place = event_detail_float.find('div', class_='place')
            if place:
                location = place.get_text(strip=True)
        
        # Extract YouTube link from description
        event_detail_wrap = soup.find('div', class_='event-detail-wrap')
        if event_detail_wrap:
            description_wrap = event_detail_wrap.find('div', class_='description-wrap')
            if description_wrap:
                # Look for YouTube link
                youtube_match = re.search(r'https://(?:www\.)?youtu\.be/[^\s<"]*', description_wrap.get_text())
                if youtube_match:
                    youtube_link = youtube_match.group(0)
        
        return speaker, location, youtube_link
    except Exception as e:
        # Silently skip errors when fetching detail pages
        return None, None, None


def parse_rss_feed(rss_content):
    """Parse RSS feed XML content."""
    try:
        root = ET.fromstring(rss_content)
        items = []
        
        # Define namespaces
        namespaces = {
            'ev': 'http://purl.org/rss/1.0/modules/event/',
            'media': 'http://search.yahoo.com/mrss/'
        }
        
        for item in root.findall('.//item'):
            title = item.find('title')
            link = item.find('link')
            guid = item.find('guid')
            description = item.find('description')
            category = item.find('category')
            pubDate = item.find('pubDate')
            
            # Get event-specific fields
            startdate = item.find('ev:startdate', namespaces)
            enddate = item.find('ev:enddate', namespaces)
            location = item.find('ev:location', namespaces)
            organizer = item.find('ev:organizer', namespaces)
            event_type = item.find('ev:type', namespaces)
            
            # Extract GUID for URL generation
            guid_text = guid.text if guid is not None else ''
            guid_number = guid_text.split('@')[0] if '@' in guid_text else ''
            
            # Generate proper event URL from GUID
            event_url = f'https://lsa.umich.edu/physics/news-events/all-events.detail.html/{guid_number}.html' if guid_number else ''
            
            event = {
                'title': title.text if title is not None else 'Untitled',
                'link': event_url,
                'guid': guid_number,
                'description': description.text if description is not None else '',
                'category': category.text if category is not None else '',
                'pubDate': pubDate.text if pubDate is not None else '',
                'startdate': startdate.text if startdate is not None else '',
                'enddate': enddate.text if enddate is not None else '',
                'location': location.text if location is not None else '',
                'organizer': organizer.text if organizer is not None else '',
                'event_type': event_type.text if event_type is not None else '',
                'speaker': None,  # Will be populated from detail page
                'detail_location': None,  # Will be populated from detail page
                'youtube_link': None,  # Will be populated from detail page
            }
            items.append(event)
        
        return items
    except ET.ParseError as e:
        print(f"Error parsing RSS feed: {e}")
        sys.exit(1)


def extract_time_from_title(title):
    """Extract time from event title (e.g., '12:00pm')."""
    # Pattern for times like "12:00pm" or "11:00am"
    match = re.search(r'(\d{1,2}:\d{2}(?:am|pm))', title, re.IGNORECASE)
    return match.group(1) if match else None


def extract_end_time_from_title(title, start_time):
    """Extract end time from title if available."""
    # Look for pattern like "12:00-1:00 PM" or "12:00pm-1:00pm"
    match = re.search(r'(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})\s*(am|pm)?', title, re.IGNORECASE)
    if match:
        return match.group(2) + (match.group(3) or 'pm')
    return None


def format_time_range(event):
    """Format time range for display."""
    start_date = parse_iso_datetime(event['startdate'])
    end_date = parse_iso_datetime(event['enddate'])
    
    if not start_date:
        return ''
    
    start_time = start_date.strftime('%I:%M %p').lstrip('0')
    
    if end_date:
        end_time = end_date.strftime('%I:%M %p').lstrip('0')
        return f"{start_time}-{end_time}"
    
    return start_time


def clean_html_description(description):
    """Clean HTML/CDATA from description."""
    # Remove CDATA tags if present
    description = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', description, flags=re.DOTALL)
    # Remove HTML tags but keep some structure
    description = re.sub(r'<[^>]+>', '', description)
    return description.strip()


def clean_event_title(title):
    """Remove parenthetical date and time from event title."""
    # Remove pattern like " (December 17, 2025 12:00pm)" at the end
    cleaned = re.sub(r'\s*\([A-Za-z]+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}(?:am|pm)\)\s*$', '', title, flags=re.IGNORECASE)
    return cleaned.strip()


def extract_location_from_description(description):
    """Extract in-person location from description if available."""
    cleaned = clean_html_description(description)
    
    # Look for "In-person:" or "In-Person:" pattern
    match = re.search(r'In-[Pp]erson:\s*(.+?)(?:$|\n|Zoom)', cleaned, re.MULTILINE)
    if match:
        location = match.group(1).strip()
        # Remove any trailing commas or extra whitespace
        location = re.sub(r',\s*$', '', location).strip()
        # If it contains extra details (address, room number), just get the main location
        if ',' in location:
            location = location.split(',')[0].strip()
        return location if location else None
    
    return None


def titlecase(text):
    """Convert text to title case, preserving articles and small words."""
    small_words = {'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'if', 'in', 'into', 'is', 'it', 'nor', 'of', 'on', 'or', 'such', 'that', 'the', 'to', 'up', 'with'}
    
    words = text.split()
    result = []
    
    for i, word in enumerate(words):
        # First and last words are always capitalized
        if i == 0 or i == len(words) - 1:
            result.append(word.capitalize())
        elif word.lower() not in small_words:
            result.append(word.capitalize())
        else:
            result.append(word.lower())
    
    return ' '.join(result)


def extract_speaker_from_description(description, title):
    """Extract speaker/organizer from description."""
    # Look for common patterns in description
    cleaned = clean_html_description(description)
    
    # Remove "Event Begins:" line as it's redundant
    cleaned = re.sub(r'^.*?Event Begins:.*?$\n?', '', cleaned, flags=re.MULTILINE)
    # Remove "Location:" line as we display it separately
    cleaned = re.sub(r'^.*?Location:.*?$\n?', '', cleaned, flags=re.MULTILINE)
    # Remove "Organized By:" line
    cleaned = re.sub(r'^.*?Organized By:.*?$\n?', '', cleaned, flags=re.MULTILINE)
    
    # Pattern: "Name (Institution/Affiliation)"
    match = re.search(r'^(.+?)\s*\(([^)]+)\)\s*$', cleaned.split('\n')[0], re.MULTILINE)
    if match:
        return match.group(1).strip()
    
    # If not found, try to extract from first line
    first_line = cleaned.split('\n')[0].strip()
    if first_line and len(first_line) < 200:  # Reasonable speaker name length
        return first_line
    
    return ''


def generate_html_output(events, start_date, end_date, output_dir):
    """Generate HTML output grouped by date."""
    # Filter events by date range and remove duplicates
    filtered_events = []
    seen_guids = set()
    
    for event in events:
        event_date = parse_iso_datetime(event['startdate'])
        if event_date and start_date <= event_date <= end_date:
            # Skip duplicate events (same GUID)
            guid = event.get('guid', '')
            if guid and guid in seen_guids:
                continue
            if guid:
                seen_guids.add(guid)
            filtered_events.append(event)
    
    # Sort by start date and time
    filtered_events.sort(key=lambda e: parse_iso_datetime(e['startdate']) or datetime.min)
    
    # Group by date
    events_by_date = defaultdict(list)
    for event in filtered_events:
        event_date = parse_iso_datetime(event['startdate'])
        if event_date:
            date_key = event_date.strftime('%Y-%m-%d')
            events_by_date[date_key].append(event)
    
    # Sort events within each day by start time
    for date_key in events_by_date:
        events_by_date[date_key].sort(key=lambda e: parse_iso_datetime(e['startdate']) or datetime.min)
    
    # Format date range for title
    start_display = start_date.strftime('%m/%d/%Y')
    end_display = end_date.strftime('%m/%d/%Y')
    title = f"Physics Seminars & Colloquia | {start_display} - {end_display}"
    
    # Build HTML
    html_parts = [
        '<!DOCTYPE html>',
        '<html>',
        '<head><meta charset="UTF-8"><title>' + title + '</title></head>',
        '<body style="font-family: Arial, sans-serif; font-size: 10pt; color: black;">',
        f'<h1>{title}</h1>',
    ]
    
    for date_key in sorted(events_by_date.keys()):
        event_date = datetime.strptime(date_key, '%Y-%m-%d')
        date_display = event_date.strftime('%A, %B %d, %Y')
        
        html_parts.append(
            f'<div style="margin-bottom: 5px;">'
            f'<a href="https://lsa.umich.edu/physics/news-events/all-events.html#date={date_key}&view=day" '
            f'style="color: #0b769f; font-weight: bold; text-decoration: underline;" target="_blank">{date_display}</a>'
            f'</div>'
        )
        
        for event in events_by_date[date_key]:
            time_range = format_time_range(event)
            title_text = titlecase(clean_event_title(event['title']))
            link = event['link']
            
            # Use detail location if available, otherwise fall back to base location
            location = event['detail_location'] or extract_location_from_description(event['description']) or event['location']
            
            speaker = event['speaker']
            youtube_link = event['youtube_link']
            
            # Find zoom link in description
            description = clean_html_description(event['description'])
            zoom_match = re.search(r'https://[^\s<"]*zoom\.us[^\s<"]*', description, re.IGNORECASE)
            zoom_url = zoom_match.group(0) if zoom_match else None
            
            html_parts.append(
                f'<div style="margin-bottom: 20px;">'
                f'<div>{time_range}</div>'
                f'<div style="font-weight: bold;"><a href="{link}" '
                f'style="color: black; text-decoration: underline;" target="_blank">{title_text}</a></div>'
            )
            
            if speaker:
                html_parts.append(f'<div style="font-style: italic;">{speaker}</div>')
            
            if location:
                html_parts.append(f'<div style="font-weight: bold;">{location}</div>')
            
            if zoom_url:
                html_parts.append(
                    f'<div style="">Event will be on Zoom: <a href="{zoom_url}" target="_blank">{zoom_url}</a></div>'
                )
            
            if youtube_link:
                html_parts.append(
                    f'<div style="">The event will be livestreamed: <a href="{youtube_link}" target="_blank">{youtube_link}</a></div>'
                )
            
            html_parts.append('</div>')
    
    html_parts.extend([
        '</body>',
        '</html>'
    ])
    
    return '\n'.join(html_parts)


def generate_google_calendar_csv(events, start_date, end_date, csv_path):
    """Generate CSV file for Google Calendar import."""
    # Filter events by date range and remove duplicates
    filtered_events = []
    seen_guids = set()
    
    for event in events:
        event_date = parse_iso_datetime(event['startdate'])
        if event_date and start_date <= event_date <= end_date:
            guid = event.get('guid', '')
            if guid and guid in seen_guids:
                continue
            if guid:
                seen_guids.add(guid)
            filtered_events.append(event)
    
    # Sort by start date
    filtered_events.sort(key=lambda e: parse_iso_datetime(e['startdate']) or datetime.min)
    
    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'Subject',
            'Start Date',
            'Start Time',
            'End Date',
            'End Time',
            'Description',
            'Location',
            'Speaker'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        for event in filtered_events:
            start_dt = parse_iso_datetime(event['startdate'])
            end_dt = parse_iso_datetime(event['enddate'])
            
            if not start_dt:
                continue
            
            title_text = titlecase(clean_event_title(event['title']))
            location = event['detail_location'] or extract_location_from_description(event['description']) or event['location']
            speaker = event['speaker'] or ''
            
            # Build description
            description_parts = []
            if speaker:
                description_parts.append(f"Speaker: {speaker}")
            if event['link']:
                description_parts.append(f"Event page: {event['link']}")
            
            # Find zoom link
            zoom_match = re.search(r'https://[^\s<"]*zoom\.us[^\s<"]*', clean_html_description(event['description']), re.IGNORECASE)
            if zoom_match:
                description_parts.append(f"Zoom: {zoom_match.group(0)}")
            
            # Find youtube link
            if event['youtube_link']:
                description_parts.append(f"YouTube: {event['youtube_link']}")
            
            description = '\n'.join(description_parts)
            
            writer.writerow({
                'Subject': title_text,
                'Start Date': start_dt.strftime('%m/%d/%Y'),
                'Start Time': start_dt.strftime('%I:%M %p'),
                'End Date': end_dt.strftime('%m/%d/%Y') if end_dt else start_dt.strftime('%m/%d/%Y'),
                'End Time': end_dt.strftime('%I:%M %p') if end_dt else '',
                'Description': description,
                'Location': location or '',
                'Speaker': speaker
            })


def sanitize_filename(name, max_length=200):
    """Return a filesystem-safe filename by replacing invalid characters.

    This targets characters invalid on Windows (<>:"/\\|?*) and trims
    trailing spaces and dots which Windows may reject. It also limits
    total length to `max_length` characters to avoid issues on some filesystems.
    """
    # Replace invalid characters with a hyphen
    safe = re.sub(r'[<>:\\"/\\|?*]', '-', name)
    # Replace multiple spaces or hyphens with a single hyphen
    safe = re.sub(r'[\s\-]+', ' ', safe).strip()
    # Replace spaces with single spaces, then convert spaces to hyphens for filenames
    safe = safe.replace(' ', ' ')
    # Trim trailing dots and spaces
    safe = safe.rstrip(' .')
    # Truncate if too long
    if len(safe) > max_length:
        safe = safe[:max_length]
    return safe


def main():
    # Feed IDs to iterate through
    feed_ids = [1965, 1178, 3798, 3799, 3767, 3801, 3811, 3247, 3804, 3805, 3806, 3807, 3813, 4897, 3606, 5034]
    
    # Aggregate events from all feeds
    all_events = []
    
    print(f"Fetching and parsing {len(feed_ids)} RSS feeds...\n")
    for feed_id in feed_ids:
        url = f"https://events.umich.edu/group/{feed_id}/rss?v=2&html_output=true"
        try:
            print(f"Fetching feed ID {feed_id}...")
            rss_content = fetch_rss_feed(url)
            events = parse_rss_feed(rss_content)
            print(f"  Found {len(events)} events")
            all_events.extend(events)
        except SystemExit:
            print(f"  Error fetching feed ID {feed_id}, skipping...")
            continue
    
    print(f"\nTotal events found (raw): {len(all_events)}")

    # Deduplicate events by URL to avoid fetching the same event page twice.
    # Keep the first occurrence of each unique `link` value. If `link` is
    # empty we do not consider it for deduplication (those remain as-is).
    deduped_events = []
    seen_links = set()
    for ev in all_events:
        link = ev.get('link') or ''
        if link:
            if link in seen_links:
                continue
            seen_links.add(link)
        deduped_events.append(ev)

    all_events = deduped_events
    print(f"Total events after deduplication by URL: {len(all_events)}")
    
    # Get date range from user
    print("\nEnter start date (m/d/yy format):")
    start_date = parse_date_input(input())
    
    print("Enter end date (m/d/yy format):")
    end_date = parse_date_input(input())
    
    if start_date > end_date:
        print("Error: Start date must be before end date")
        sys.exit(1)
    
    # Filter events within date range first
    events_in_range = []
    for event in all_events:
        event_date = parse_iso_datetime(event['startdate'])
        if event_date and start_date <= event_date <= end_date:
            events_in_range.append(event)
    
    print(f"Events within date range: {len(events_in_range)}")
    
    # Fetch detail pages only for events within the date range
    print("\nFetching event details...")
    for i, event in enumerate(events_in_range):
        if event['link']:
            print(f"  Fetching details for event {i+1}/{len(events_in_range)}: {event['title'][:50]}...")
            speaker, detail_location, youtube_link = fetch_event_detail_page(event['link'])
            event['speaker'] = speaker
            event['detail_location'] = detail_location
            event['youtube_link'] = youtube_link
    
    # Create output directory
    output_dir = 'Physics Seminars & Colloquia'
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate HTML
    print("Generating HTML output...")
    html_output = generate_html_output(all_events, start_date, end_date, output_dir)
    
    # Format date range for filenames
    start_display = start_date.strftime('%m-%d-%Y')
    end_display = end_date.strftime('%m-%d-%Y')
    
    # Save HTML file with date range in filename (sanitize for Windows)
    raw_html_filename = f"Physics Seminars & Colloquia | {start_display} - {end_display}.html"
    html_filename = sanitize_filename(raw_html_filename)
    html_path = os.path.join(output_dir, html_filename)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_output)
    
    print(f"HTML output saved to {html_path}")
    
    # Generate Google Calendar CSV
    print("Generating Google Calendar CSV...")
    raw_csv_filename = f"Physics Seminars & Colloquia | {start_display} - {end_display}.csv"
    csv_filename = sanitize_filename(raw_csv_filename)
    csv_path = os.path.join(output_dir, csv_filename)
    generate_google_calendar_csv(all_events, start_date, end_date, csv_path)
    print(f"CSV output saved to {csv_path}")


if __name__ == '__main__':
    main()
