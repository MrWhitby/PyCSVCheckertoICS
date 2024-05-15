import csv
from datetime import datetime


def determine_date_format(date_str):
    """ Try to determine the format of the date string based on expected formats """
    # List of date formats to check against
    formats = ["%b %d %Y", "%d/%m/%Y", "%m/%d/%Y"]
    for fmt in formats:
        try:
            datetime.strptime(date_str, fmt)
            return fmt
        except ValueError:
            continue
    raise ValueError("Date format not recognized")


def create_ics_from_csv(csv_path, ics_path):
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        date_format = None  # To store the detected date format

        events = []
        for row in reader:
            if not date_format:
                # Determine the date format from the first row
                date_format = determine_date_format(row['Domain expiration date'])

            # Check the current row's date format
            current_format = determine_date_format(row['Domain expiration date'])
            if current_format != date_format:
                raise ValueError("Inconsistent date formats detected")

            domain = row['Domain Name']
            expiration_date = datetime.strptime(row['Domain expiration date'], date_format)
            formatted_date = expiration_date.strftime("%Y%m%d")

            # Build event details
            event_details = {
                "domain": domain,
                "date": formatted_date
            }
            events.append(event_details)

        # Write to ICS file after checking all dates
        with open(ics_path, 'w', encoding='utf-8') as icsfile:
            icsfile.write("BEGIN:VCALENDAR\n")
            icsfile.write("VERSION:2.0\n")
            icsfile.write("PRODID:-//Your Company//Your Product//EN\n")
            for event in events:
                icsfile.write("BEGIN:VEVENT\n")
                icsfile.write(f"SUMMARY:Domain expiration: {event['domain']}\n")
                icsfile.write(f"DTSTART;VALUE=DATE:{event['date']}\n")
                icsfile.write(f"DTEND;VALUE=DATE:{event['date']}\n")
                icsfile.write(f"DESCRIPTION:Expiration date for domain {event['domain']}.\n")
                icsfile.write("RRULE:FREQ=YEARLY\n")
                icsfile.write("BEGIN:VALARM\n")
                icsfile.write("TRIGGER:-P1D\n")
                icsfile.write("ACTION:DISPLAY\n")
                icsfile.write("DESCRIPTION:Reminder\n")
                icsfile.write("END:VALARM\n")
                icsfile.write("END:VEVENT\n")
            icsfile.write("END:VCALENDAR\n")


# Example usage
csv_file_path = 'C:\\Users\\jacob\\Downloads\\Domain_List.csv'  # Adjusted path
ics_file_path = 'C:\\Users\\jacob\\Downloads\\domain_expirations.ics'  # Output file

create_ics_from_csv(csv_file_path, ics_file_path)
