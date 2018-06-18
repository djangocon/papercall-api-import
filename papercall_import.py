import frontmatter

from envparse import env, ConfigurationError
from os import makedirs
from requests import get
from slugify import slugify
from xlwt import easyxf, Workbook

# Possible submission states
SUBMISSION_STATES = ('submitted', 'accepted', 'rejected', 'waitlist')

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
    """
    Get file name from user.
    """
    output_filename = input(input_text) or default_filename

    return output_filename


def create_excel(api_key, xls_file):
    """
    Creates an Excel workbook with a spreadsheet for each status of submission.
    """
    total_submissions = 0
    total_ratings = 0
    total_feedback = 0

    # Create the Spreadsheet Workbook
    wb = Workbook()

    for submission_state in SUBMISSION_STATES:
        # Reset row counter for the new sheet
        # Row 0 is reserved for the header
        num_row = 1

        # Create the new sheet and header row for each talk state
        ws = wb.add_sheet(submission_state.upper())
        ws.write(0, 0, 'ID', HEADER_STYLE)
        ws.write(0, 1, 'Title', HEADER_STYLE)
        ws.write(0, 2, 'Format', HEADER_STYLE)
        ws.write(0, 3, 'Audience', HEADER_STYLE)
        ws.write(0, 4, 'Rating', HEADER_STYLE)
        ws.write(0, 5, 'Name', HEADER_STYLE)
        ws.write(0, 6, 'Email', HEADER_STYLE)
        ws.write(0, 7, 'Bio', HEADER_STYLE)

        for x in range(8, 35):
            ws.write(0, x, 'Comments / Feedback {}'.format(x - 7), HEADER_STYLE)

        r = get(
            'https://www.papercall.io/api/v1/submissions?_token={0}&state={1}&per_page=1000'.format(
                api_key,
                submission_state,
            )
        )

        for submission in r.json():
            total_submissions += 1
            ws.write(num_row, 0, submission['id'])
            ws.write(num_row, 1, submission['talk']['title'])
            ws.write(num_row, 2, submission['talk']['talk_format'])
            ws.write(num_row, 3, submission['talk']['audience_level'])
            ws.write(num_row, 4, submission['rating'])

            if 'profile' in submission:
                ws.write(num_row, 5, submission['profile']['name'])
                ws.write(num_row, 6, submission['profile']['email'])
                ws.write(num_row, 7, submission['profile']['bio'])
            else:
                ws.write(num_row, 5, 'Not Revealed')
                ws.write(num_row, 6, 'Not Revealed')
                ws.write(num_row, 7, 'Not Revealed')

            # Start at column 8 for comments and feedback
            num_col = 8

            # Only include ratings comments if they've been entered
            c = get(
                'https://www.papercall.io/api/v1/submissions/{}/ratings?_token={}'.format(
                    submission['id'],
                    api_key,
                )
            )
            for ratings_comment in c.json():
                total_ratings += 1

                if len(ratings_comment['comments']):
                    ws.write(
                        num_row,
                        num_col,
                        '(Comment from {}) {}'.format(
                            ratings_comment['user']['email'],
                            ratings_comment['comments'],
                        ),
                    )
                    num_col += 1

            # Loop through all of the submitter / reviewer feedback and include after comments
            f = get(
                'https://www.papercall.io/api/v1/submissions/{}/feedback?_token={}'.format(
                    submission['id'],
                    api_key,
                )
            )
            for feedback in f.json():
                total_feedback += 1
                ws.write(
                    num_row,
                    num_col,
                    '(Feedback from {}) {}'.format(
                        feedback['user']['email'],
                        feedback['body'],
                    ),
                )
                num_col += 1

            num_row += 1

    wb.save(xls_file)

    return total_submissions, total_ratings, total_feedback


def create_yaml(api_key, yaml_dir):
    for submission_state in SUBMISSION_STATES:
        # Create the directories, if they don't exist.
        makedirs(
            '{}/{}'.format(
                yaml_dir,
                submission_state,
            ), exist_ok=True,
        )

        r = get(
            'https://www.papercall.io/api/v1/submissions?_token={0}&state={1}&per_page=1000'.format(
                api_key,
                submission_state,
            )
        )

        for submission in r.json():
            talk_format = None
            if submission['talk']['talk_format'][0:4].lower() == "talk":
                talk_format = "talk"
            elif submission['talk']['talk_format'][0:4].lower() == "tuto":
                talk_format = "tutorial"

            if talk_format:
                talk_title_slug = slugify(submission['talk']['title'])

                post = frontmatter.loads(submission['talk']['description'])
                post['abstract'] = submission['talk']['abstract']
                post['category'] = talk_format
                post['title'] = submission['talk']['title']
                post['difficulty'] = submission['talk']['audience_level']
                post['permalink'] = '/{}/{}/'.format(
                    talk_format,
                    talk_title_slug,
                )
                post['layout'] = 'session-details'
                post['accepted'] = True if submission_state == 'accepted' else False
                post['published'] = True
                post['sitemap'] = True

                # TODO: Scheduling info...
                post['date'] = '2016-07-18 09:00'
                post['room'] = ''
                post['track'] = ''

                # TODO: Determine if we still need summary (I don't think we do)
                post['summary'] = ''

                # todo: refactor template layout to support multiple authors
                post['presenters'] = [
                    {
                        'name': submission['profile']['name'],
                        'bio': submission['profile']['bio'],
                        'company': submission['profile']['company'],
                        'photo_url': '',
                        'github': '',
                        'twitter': submission['profile']['twitter'],
                        'website': submission['profile']['url'],
                    },
                ]

                # post conference info
                post['video_url'] = ''
                post['slides_url'] = ''

                with open(
                    '{}/{}/{}-{}.md'.format(
                        yaml_dir,
                        submission_state,
                        talk_format,
                        talk_title_slug,
                    ),
                    'wb'
                ) as file_to_write:
                    frontmatter.dump(post, file_to_write)

                print(frontmatter.dumps(post))


def main():
    try:
        api_key = env('PAPERCALL_API_KEY')
    except ConfigurationError:
        api_key = get_api_key()
    file_format = get_format()

    if file_format == "1":
        xls_file = get_filename('Filename to write [djangoconus.xls]: ', 'djangoconus.xls')
        total_submissions, total_ratings, total_feedback = create_excel(api_key, xls_file)

        print(
            """
    Total Submissions: {total_submissions}
    Total Ratings: {total_ratings}
    Total Feedback: {total_feedback}
    """.format(
                total_submissions=total_submissions,
                total_ratings=total_ratings,
                total_feedback=total_feedback,
            )
        )
    elif file_format == "2":
        yaml_dir = get_filename('Directory to write to [yaml]: ', 'yaml')
        create_yaml(api_key, yaml_dir)


if __name__ == "__main__":
    main()
