from enum import Enum
from datetime import datetime
import docx2txt
import re

NEWLINE = '\n'
TBC = 'TBC'
TITLES_TO_GENDER = {
    'Master': 'Male',
    'Mr': 'Male',
    'Miss': 'Female',
    'Mrs': 'Female',
    'Ms': 'Female',
    'Mx': 'Non-Binary'
}
PLAN_MANAGED_EMAIL = 'planmanaged@email.com'
NDIA_MANAGED_EMAIL = 'michelle@lightstreetcare.com.au'
MAX_32_BIT_INT = 2147483647


class SupportsType(Enum):
    CORE = 1
    CAPACITY_BUILDING = 2
    CAPITAL = 3


class Location:
    def __init__(self, address):
        try:
            # Get the house number
            end = index(address, ' ')[0]
            self.house_number = address[:end]

            # Get the street
            start = end + 1
            end = index(address, '^(?:[^ ]* ){2}', start)[1] - 1
            self.street = address[start:end].title()

            # Get the suburb
            start = end + 1
            end = index(address, r' .* \d{4}$', start)[0]
            self.suburb = address[start:end].title()

            # Get the state
            start = end + 1
            end = index(address, ' ', start)[0]
            self.state = address[start:end].upper()

            # Get the postcode
            start = index(address, r'\d{4}$')[0]
            self.postcode = address[start:]
        except TypeError:
            self.house_number = ''
            self.street = ''
            self.suburb = ''
            self.state = ''
            self.postcode = ''

    def __str__(self):
        return (
            f'{self.house_number} '
            f'{self.street} '
            f'{self.suburb} '
            f'{self.state} '
            f'{self.postcode}'
        )


class Client:
    def __init__(self,
                 title,
                 full_name,
                 gender,
                 dob,
                 address,
                 home_phone_number,
                 mobile_phone_number,
                 email_address,
                 ndis_number):
        self.title = title
        self.full_name = full_name
        self.gender = gender
        self.dob = dob
        self.address = address
        self.home_phone_number = home_phone_number
        self.mobile_phone_number = mobile_phone_number
        self.email_address = email_address
        self.ndis_number = ndis_number

        # Get the first name
        end = index(full_name, ' ')[0]
        self.first_name = full_name[:end]

        # Get the last name
        start = end + 1
        self.last_name = full_name[start:]


class Plan:
    def __init__(self, start_date, end_date):
        self.start_date = start_date
        self.end_date = end_date


class Supports:
    def __init__(self, goals, categories, total):
        self.goals = goals
        self.categories = categories
        self.total = total


class Record:
    def __init__(self,
                 client,
                 plan,
                 supports,
                 support_coordination_management_type,
                 support_coordination_hours,
                 additional_email_address,
                 service_region_id):
        self.client = client
        self.plan = plan
        self.supports = supports
        self.support_coordination_management_type = support_coordination_management_type
        self.support_coordination_hours = support_coordination_hours
        self.additional_email_address = additional_email_address
        self.service_region_id = service_region_id

    def __str__(self):
        string = (
            f'Client Information:\n'
            f'    Title: {self.client.title}\n'
            f'    First Name: {self.client.first_name}\n'
            f'    Last Name: {self.client.last_name}\n'
            f'    Gender: {self.client.gender}\n'
            f'    Date of Birth: {self.client.dob}\n'
            f'    Address: {str(self.client.address)}\n'
            f'    Home Phone Number: {self.client.home_phone_number}\n'
            f'    Mobile Phone Number: {self.client.mobile_phone_number}\n'
            f'    Email Address: {self.client.email_address}\n'
            f'    NDIS Number: {self.client.ndis_number}\n\n'

            f'Plan:\n'
            f'    Start Date: {self.plan.start_date}\n'
            f'    End Date: {self.plan.end_date}\n\n'

            f'Supports:\n'
        )

        # Add supports data to the string
        for section, supports in self.supports.items():
            string += f'    {section}:\n'

            string += '        Goals:'
            if supports.goals == TBC:
                string += f' {TBC}\n'
            else:
                string += '\n'
                for goal in supports.goals:
                    string += f'            - {goal}\n'

            string += '        Categories:'
            if supports.categories == TBC:
                string += f' {TBC}\n'
            else:
                string += '\n'
                for category in supports.categories:
                    string += f'            {category[0]}: {category[1]}\n'

            string += f'        Total: {supports.total}\n'
        string += '\n'

        string += (
            f'Additional Information:\n'
            f'    Support Coordination:\n'
            f'        Mangement Type: {self.support_coordination_management_type}\n'
            f'        Hours: {self.support_coordination_hours}\n'
            f'    Additional Email Address: {self.additional_email_address}\n'
            f'    Service Region ID: {self.service_region_id}'
        )

        return string


def clean_document(document):
    """Cleans a document by removing common inconsistencies

    Args:
        document (str): The document to clean

    Returns:
        str: The cleaned document

    """
    doc = re.sub(' {2,}', ' ', document)
    doc = re.sub(r'\n{2,}', r'\n', doc)
    doc = re.sub('to to', 'to', doc)

    return doc


def get_document(path):
    """Gets the contents of a word document

    Args:
        path (str): The path to a word document

    Returns:
        str: The contents of the word document

    """
    return clean_document(docx2txt.process(path))


def index(string, regex, start=0):
    """Get the start and end indicies of a found regex pattern in a string

    Args:
        string (str): The The contents of a document
        regex (str): The regex pattern to search for
        start (int): The index to start the search from (optional)

    Returns:
        (int, int): A 2-tuple containing the start and end index of the text found in a string
            that matches the regex pattern, or None if the regex pattern couldn't be found

    """
    match = re.search(regex, string[start:], re.IGNORECASE)
    if match is not None:
        return tuple(index + start for index in match.span())


def clean_string(string):
    """Cleans a string by removing all whitespace characters (space, tab, newline, etc.)

    Args:
        string (str): The string to clean

    Returns:
        str: The cleaned string

    """
    return ' '.join(string.split())


def get_title(document):
    """Extracts a title out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted title, or 'TBC' if it could not be found

    """
    try:
        start = index(document, r'reference.*\n')[1]
        end = index(document, r'( |\.)', start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_full_name(document):
    """Extracts a full name out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted full name, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'name: ')[1]
        end = index(document, 'ndis', start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end]).title()


def get_dob(document):
    """Extracts a date of birth out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted date of birth, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'date of birth')[1]
        start = index(document, r'\d', start)[0]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return datetime.strptime(clean_string(document[start:end]), '%d %B %Y').strftime('%d/%m/%Y')


def get_address(document):
    """Extracts an address out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted address, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'reference.*')[1]
        start = index(document, r'\d', start)[0]
        end = index(document, r'\d{4}\n', start)[1]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_ndis_number(document):
    """Extracts an NDIS number out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted NDIS number, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'ndis number: ')[1]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_plan_start_date(document):
    """Extracts a plan start date out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted plan start date, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'start date: ')[1]
        end = index(document, 'ndis', start)[0]
    except TypeError:
        return TBC

    return datetime.strptime(clean_string(document[start:end]), '%d %B %Y').strftime('%d/%m/%Y')


def get_plan_end_date(document):
    """Extracts a plan end date out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted plan end date, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'review due date: ')[1]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return datetime.strptime(clean_string(document[start:end]), '%d %B %Y').strftime('%d/%m/%Y')


def get_home_phone_number(document):
    """Extracts a home phone number out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted home phone number, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'home number: ')[1]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_mobile_phone_number(document):
    """Extracts a mobile phone number out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted mobile phone number, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'mobile: ')[1]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_email_address(document):
    """Extracts an email address out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted email address, or 'TBC' if it could not be found

    """
    try:
        start = index(document, r'preferred contact method.*email\n')[1]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_core_supports_included_funding(document):
    """Extracts the core supports included funding out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted core supports included funding, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'core supports')[0]
        start = index(document, 'funding for', start)[0]
        end = index(document, r'[a-z]\.', start)[1] + 1
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_support_coordination_management_type(document):
    """Extracts the support coordination management type out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted support coordination management type, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'support coordination')[1]
        start = index(document, 'self-managed|plan-managed|ndia-managed', start)[0]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_supports_goals(document, supports_section):
    """Extracts supports goals out of a document

    Args:
        document (str): The contents of a document
        supports_section (SupportsType): The supports category to search for goals

    Returns:
        tuple(str): The extracted supports goals, or 'TBC' if none could be found

    """
    if supports_section == SupportsType.CORE:
        regex_strings = (
            'goal/s my core supports',
            'core supports'
        )
    elif supports_section == SupportsType.CAPACITY_BUILDING:
        regex_strings = (
            'goal/s my capacity building supports',
            'capacity building funding'
        )
    elif supports_section == SupportsType.CAPITAL:
        regex_strings = (
            'goal/s my capital supports',
            'capital supports funding'
        )
    else:
        return

    goals = []
    try:
        start = index(document, regex_strings[0])[0]
        start = index(document, NEWLINE, start)[0] + 1
        end = index(document, NEWLINE, start)[0]

        goal = document[start:end]
        while regex_strings[1] not in goal.lower():
            goals.append(goal)

            start = end + 1
            end = index(document, NEWLINE, start)[0]
            goal = document[start:end]
    except TypeError:
        return TBC

    return tuple(goals)


def get_supports_categories(document, supports_section):
    """Extracts supports categories and their budgets out of a document

    Args:
        document (str): The contents of a document
        supports_section (SupportsType): The supports section to search for supports

    Returns:
        tuple(tuple(str, str)): The extracted supports categories and their budgets,
            or 'TBC' if none could be found

    """
    budget_regex = r'\$.*\.\d{2}\n'
    if supports_section == SupportsType.CORE:
        categories_to_budgets = [['Core']]
        categories = [
            'Assistance with Daily Life',
            'Transport',
            'Consumables',
            'Assistance with Social, Economic and Community Participation'
        ]
        regex_string = 'core supports'
    elif supports_section == SupportsType.CAPACITY_BUILDING:
        categories_to_budgets = []
        categories = [
            'Support Coordination',
            'Improved Living Arrangements',
            'Increased Social and Community Participation',
            'Finding and Keeping a Job',
            'Improved Relationships',
            'Improved Health and Wellbeing',
            'Improved Learning',
            'Improved Life Choices',
            'Improved Daily Living'
        ]
        regex_string = 'capacity building supports'
    elif supports_section == SupportsType.CAPITAL:
        categories_to_budgets = []
        categories = [
            'Assistive Technology',
            'Home Modifications and Specialist Disability Accommodation'
        ]
        regex_string = 'capital supports'
    else:
        return

    try:
        category_start = index(document, regex_string)[1]

        # Get all the applicable core supports categories
        for i in range(len(categories)):
            lowest_index = MAX_32_BIT_INT
            curr_category = None
            for category in categories:
                curr_indices = index(document, category + r'(?!.*\.)', category_start)
                if curr_indices is None:
                    continue

                # Find the closest category in the document
                if curr_indices[0] < lowest_index:
                    lowest_index = curr_indices[0]
                    curr_category = category

            # Add the correct category to the list, if possible
            if curr_category is not None:
                categories_to_budgets.append([curr_category])
                categories.pop(categories.index(curr_category))

        # Add budgets to the list for each category
        indices = index(document, budget_regex, category_start)
        for i in range(len(categories_to_budgets)):
            categories_to_budgets[i].append(document[indices[0]:indices[1] - 1])
            indices = index(document, budget_regex, indices[1])

    except TypeError:
        return TBC

    return tuple(tuple(elem) for elem in categories_to_budgets)


def get_supports_total(document, supports_section):
    """Extracts supports total budgets from a document

    Args:
        document (str): The contents of a document
        supports_section (SupportsType): The supports section to get the total budget from

    Returns:
        str: The extracted supports toal budget, or 'TBC' if it could not be found

    """
    if supports_section == SupportsType.CORE:
        regex_string = 'total core supports'
    elif supports_section == SupportsType.CAPACITY_BUILDING:
        regex_string = 'total capacity building supports'
    elif supports_section == SupportsType.CAPITAL:
        regex_string = 'total capital supports'
    else:
        return

    try:
        start = index(document, regex_string)[0]
        start = index(document, r'\$.*\n[^\$]', start)[0]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def get_funded_supports_total(document):
    """Extracts a plan start date out of a document

    Args:
        document (str): The contents of a document

    Returns:
        str: The extracted plan start date, or 'TBC' if it could not be found

    """
    try:
        start = index(document, 'total funded supports')[0]
        start = index(document, r'\$', start)[0]
        end = index(document, NEWLINE, start)[0]
    except TypeError:
        return TBC

    return clean_string(document[start:end])


def build_record_from_document(path):
    """Build a Record object from a document

    Args:
        path (str): The path to a word document

    Returns:
        Record: The built Record object

    """
    # Get the contents of the word document
    document = get_document(path)

    # Get address by building a Location object
    address = Location(get_address(document))

    # Build a Client object
    title = get_title(document)
    client = Client(
        title,
        get_full_name(document),
        TITLES_TO_GENDER.get(title),
        get_dob(document),
        address,
        get_home_phone_number(document),
        get_mobile_phone_number(document),
        get_email_address(document),
        get_ndis_number(document)
    )

    # Build a Plan object
    plan = Plan(get_plan_start_date(document), get_plan_end_date(document))

    # Build a supports dictionary
    supports = {
        'Core': Supports(
            get_supports_goals(document, SupportsType.CORE),
            get_supports_categories(document, SupportsType.CORE),
            get_supports_total(document, SupportsType.CORE)
        ),

        'Capacity Building': Supports(
            get_supports_goals(document, SupportsType.CAPACITY_BUILDING),
            get_supports_categories(document, SupportsType.CAPACITY_BUILDING),
            get_supports_total(document, SupportsType.CAPACITY_BUILDING)
        ),

        'Capital': Supports(
            get_supports_goals(document, SupportsType.CAPITAL),
            get_supports_categories(document, SupportsType.CAPITAL),
            get_supports_total(document, SupportsType.CAPITAL)
        )
    }

    # Get the support coordination management type
    support_coordination_management = get_support_coordination_management_type(document)

    # Get the additional email address
    additional_email_address = (
        PLAN_MANAGED_EMAIL
        if index(support_coordination_management, 'plan-managed') is not None
        else NDIA_MANAGED_EMAIL
    )

    # Build a Record object
    record = Record(
        client,
        plan,
        supports,
        support_coordination_management,
        TBC,
        additional_email_address,
        TBC
    )

    return record


def build_record_from_string(string):
    """Build a Record object from a Record object string

    Args:
        string (str): A Record object string

    Returns:
        Record: The built Record object, or None if the input string is invalid

    """
    try:
        lines = string.split('\n')

        # The indices for lines that have constant formatting (i.e. not lists that change size)
        const_indices = list(range(1, 11))
        const_indices.extend([13, 14])
        const_data = [lines[i][index(lines[i], ':.')[1]:].strip() for i in const_indices]

        # Line number to start on for supports data
        line_index = 18

        supports_sections = ('Core', 'Capacity Building', 'Capital')
        supports = {}
        for section in supports_sections:
            # Goals
            if TBC in lines[line_index]:
                goals = TBC
                line_index += 1
            else:
                line_index += 1
                curr_line = lines[line_index]

                # Get goals
                goals = []
                while '-' in curr_line:
                    goals.append(curr_line[index(curr_line, '-.')[1]:].strip())
                    line_index += 1
                    curr_line = lines[line_index]

            # Categories
            if TBC in lines[line_index]:
                categories = TBC
                line_index += 1
            else:
                line_index += 1
                curr_line = lines[line_index].strip()

                # Get categories
                categories = []
                while 'total' not in curr_line.lower():
                    mid = index(curr_line, ':.')
                    category = curr_line[:mid[0]]
                    budget = curr_line[mid[1]:].strip()
                    categories.append((category, budget))

                    line_index += 1
                    curr_line = lines[line_index].strip()

                categories = tuple(categories)

            # Total
            curr_line = lines[line_index]
            total = curr_line[index(curr_line, ':.')[1]:].strip()
            line_index += 2

            supports[section] = Supports(goals, categories, total)

        # Get the line index to the last section of constant data
        line_index += 2
        const_indices = list(range(line_index, line_index + 4))
        const_data.extend([lines[i][index(lines[i], ':.')[1]:].strip() for i in const_indices])

        # Build a Client object
        client = Client(
            const_data[0],
            f'{const_data[1]} {const_data[2]}',
            const_data[3],
            const_data[4],
            Location(const_data[5]),
            const_data[6],
            const_data[7],
            const_data[8],
            const_data[9]
        )

        # Build a Plan object
        plan = Plan(const_data[10], const_data[11])

        # Build a Record object
        record = Record(
            client,
            plan,
            supports,
            const_data[12],
            const_data[13],
            const_data[14],
            TBC
        )

        return record
    except TypeError:
        return None
