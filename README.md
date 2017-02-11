# PaperCall.io API to XLS

This script calls the PaperCall.io API and pulls submissions into a spreadsheet for each state (submitted, accepted, rejected, waitlist).

# Installation

Use your favorite tool to create a `virtualenv`, then:

    git clone https://github.com/djangocon/papercall-api-to-xls.git
    cd papercall-api-to-xls
    pip install requirements.txt

# Running the Script

You'll need to know your event ID number. Then get your API key from:

https://www.papercall.io/events/[event_id]/apidocs

Then run the command:

    python apitoxls.py