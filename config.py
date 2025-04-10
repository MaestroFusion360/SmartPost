import os

# Default values for the add-in settings

DEBUG = False # Set to True to enable debug mode, False to disable
ADDIN_NAME = os.path.basename(os.path.dirname(__file__))

COMPANY_NAME = ''
DEFAULT_PERSONAL_LICENSE = 'false'
DEFAULT_PROGRAM_NAME = '1001'
DEFAULT_PROGRAM_NUMBER = '1001'
DEFAULT_COMMENT = ''
DEFAULT_POST_NAME = ''
DEFAULT_POST_FOLDER = os.path.expanduser('~/AppData/Roaming/Autodesk/Fusion 360 CAM/Posts')
DEFAULT_OUTPUT_FOLDER = os.path.expanduser('D:/Desktop')
DEFAULT_UNIT = 'Document Unit'
DEFAULT_IS_OPEN_IN_EDITOR = 'false'
DEFAULT_ALLOW_HELICAL_MOVES = 'true'
DEFAULT_HIGH_FEEDRATE_MAPPING_VALUE = 'Preserve rapid movement'
DEFAULT_MINIMUM_CHORD_LENGTH = '0.1'
DEFAULT_HIGH_FEEDRATE = '0'
DEFAULT_MAXIMUM_CIRCULAR_RADIUS = '1000'
DEFAULT_MINIMUM_CIRCULAR_RADIUS = '0.01'
DEFAULT_TOLERANCE = '0.001'

# Unique palette ID
sample_palette_id = f'{COMPANY_NAME}_{ADDIN_NAME}_palette_id'
