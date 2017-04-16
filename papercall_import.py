from requests import get
import xlwt

# Possible proposal states
PROPOSAL_STATES = ('submitted', 'accepted', 'rejected', 'waitlist')

# Style for the Spreadsheet headers
HEADER_STYLE = xlwt.easyxf(
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
    print('2: Markdown')
    file_format = input('Please enter your your output format (1 or 2): ')
    if file_format not in ('1', '2'):
        raise ValueError('Error: Output format must be "1" or "2".')

    return file_format


def get_xls_file():
    # Get XLSination file name
    xls_file = input('Filename to write [djangoconus.xls]: ') or 'djangoconus.xls'

    return xls_file


def create_excel(api_key, xls_file):
    # Create the Spreadsheet Workbook
    wb = xlwt.Workbook()

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
            'https://www.papercall.io/api/v1/submissions?_token={0}&state={1}'.format(
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


def main():
    api_key = get_api_key()
    file_format = get_format()

    if file_format == "1":
        xls_file = get_xls_file()
        create_excel(api_key, xls_file)
    elif file_format == 2:
        pass


if __name__ == "__main__":
    main()
