import pytest
from win32com.client import constants


@pytest.fixture(scope="module")
def tlbs():
    from xlviews.client import iter_typelib_specs

    return list(iter_typelib_specs())


@pytest.mark.parametrize(
    "clsid",
    [
        "{00020813-0000-0000-C000-000000000046}",
        "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}",
        "{91493440-5A91-11CF-8700-00AA0060263B}",
    ],
)
def test_iter_typelib_specs(tlbs, clsid):
    clsids = [tlb.clsid for tlb in tlbs]
    assert clsid in clsids


def test_ensure_modules():
    from xlviews.client import ensure_modules

    ensure_modules()
    assert constants.msoConnectorStraight == 1
