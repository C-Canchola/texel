from .book import get_bk
from .decorators import create_active_app
from .constants import TEST_FILE_PATH


def get_test_bk(**kwargs):
    return get_bk(TEST_FILE_PATH, **kwargs)
