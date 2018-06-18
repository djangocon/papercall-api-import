# PaperCall.io API Import

This script calls the PaperCall.io API and pulls submissions into a format chosen by the user. The user can choose:

- A spreadsheet for each state (submitted, accepted, rejected, waitlist).
- A directory of YAML files with all four states and talks within.

# Installation

Use your favorite tool to create a `virtualenv`, then:

    git clone https://github.com/djangocon/papercall-api-import.git
    cd papercall-api-import
    pip install -r requirements.txt

# Running the Script

You'll need to know your event ID number. Then get your API key from:

https://www.papercall.io/events/[event_id]/apidocs

Then run the command:

    python papercall_import.py

...and follow the input prompts!
