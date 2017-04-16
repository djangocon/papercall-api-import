from os import makedirs
from requests import get
from slugify import slugify
from xlwt import easyxf, Workbook

# Possible proposal states
PROPOSAL_STATES = ('submitted', 'accepted', 'rejected', 'waitlist')

# Style for the Spreadsheet headers
HEADER_STYLE = easyxf(
    'font: name Verdana, color-index blue, bold on',
    num_format_str='#,##0.00'
)


def get_api_key():
    """
    Get the user's API key
    """
    print('Your DjangoCon PaperCall API Key can be found here: https://www.papercall.io/events/316/apidocs')
    api_key = input('Please enter your PaperCall event API Key: ')
    if len(api_key) != 32:
        raise ValueError('Error: API Key must be 32 characters long.')

    return api_key


def get_format():
    """
    Get the output format to write to.
    """
    print('Which format would you like to output?')
    print('1: Excel')
    print('2: YAML/Markdown for Jekyll')
    file_format = input('Please enter your your output format (1 or 2): ')
    if file_format not in ('1', '2'):
        raise ValueError('Error: Output format must be "1" or "2".')

    return file_format


def get_filename(input_text, default_filename):
    # Get file name from user
    output_filename = input(input_text) or default_filename

    return output_filename


def create_excel(api_key, xls_file):
    # Create the Spreadsheet Workbook
    wb = Workbook()

    for ps in PROPOSAL_STATES:
        # Reset row counter for the new sheet
        # Row 0 is reserved for the header
        num_row = 1

        # Create the new sheet and header row for each talk state
        ws = wb.add_sheet(ps.upper())
        ws.write(0, 0, 'ID', HEADER_STYLE)
        ws.write(0, 1, 'Title', HEADER_STYLE)
        ws.write(0, 2, 'Format', HEADER_STYLE)
        ws.write(0, 3, 'Audience', HEADER_STYLE)
        ws.write(0, 4, 'Rating', HEADER_STYLE)

        r = get(
            'https://www.papercall.io/api/v1/submissions?_token={0}&state={1}&per_page=1000'.format(
                api_key,
                ps,
            )
        )

        for proposal in r.json():
            ws.write(num_row, 0, proposal['id'])
            ws.write(num_row, 1, proposal['talk']['title'])
            ws.write(num_row, 2, proposal['talk']['talk_format'])
            ws.write(num_row, 3, proposal['talk']['audience_level'])
            ws.write(num_row, 4, proposal['rating'])

            num_row += 1

    wb.save(xls_file)


def create_yaml(api_key, yaml_dir):
    for ps in PROPOSAL_STATES:
        # Create the directories, if they don't exist.
        makedirs(
            '{}/{}'.format(
                yaml_dir,
                ps,
            ), exist_ok=True,
        )

        r = get(
            'https://www.papercall.io/api/v1/submissions?_token={0}&state={1}&per_page=1000'.format(
                api_key,
                ps,
            )
        )

        from pprint import pprint, pformat
        for proposal in r.json():
            talk_format = None
            if proposal['talk']['talk_format'][0:4].lower() == "talk":
                talk_format = "talk"
            elif proposal['talk']['talk_format'][0:4].lower() == "tuto":
                talk_format = "tutorial"

            if talk_format:
                with open(
                    '{}/{}/{}-{}.md'.format(
                        yaml_dir,
                        ps,
                        talk_format,
                        slugify(proposal['talk']['title']),
                    ),
                    'w'
                ) as file_to_write:
                    file_to_write.write(pformat(proposal))

                pprint(proposal)


def main():
    api_key = get_api_key()
    file_format = get_format()

    if file_format == "1":
        xls_file = get_filename('Filename to write [djangoconus.xls]: ', 'djangoconus.xls')
        create_excel(api_key, xls_file)
    elif file_format == "2":
        yaml_dir = get_filename('Directory to write to [yaml]: ', 'yaml')
        create_yaml(api_key, yaml_dir)


if __name__ == "__main__":
    main()
