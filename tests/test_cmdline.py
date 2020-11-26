import os
import io
import zipfile
import json
import shutil
import sys
import tempfile
from unittest.mock import patch

import pytest

import tnefparse.cmdline


def test_cmdline_overview(script_runner):
    ret = script_runner.run('tnefparse', '-o', 'tests/examples/body.tnef')
    assert ret.success
    assert "Overview of tests/examples/body.tnef" in ret.stdout
    assert ret.stderr == ''
    assert ret.success


def test_cmdline_attch_extract(script_runner):
    tmpdir = tempfile.mkdtemp()
    ret = script_runner.run('tnefparse', '-a', '-p', tmpdir, 'tests/examples/one-file.tnef')
    assert ret.stderr == 'Successfully wrote 1 files\n'
    assert ret.success
    shutil.rmtree(tmpdir)


def test_cmdline_zip_extract(script_runner):
    tmpdir = tempfile.mkdtemp()
    ret = script_runner.run('tnefparse', '-z', '-p', tmpdir, 'tests/examples/one-file.tnef')
    assert os.path.isfile(tmpdir + '/attachments.zip')
    assert ret.stderr == 'Successfully wrote attachments.zip\n'
    with open(tmpdir + '/attachments.zip', 'rb') as zip_fp:
        zip_stream = io.BytesIO(zip_fp.read())
    zip_file = zipfile.ZipFile(zip_stream)
    assert zip_file.namelist() == ['AUTHORS']
    assert ret.success
    shutil.rmtree(tmpdir)


def test_cmdline_no_body(script_runner):
    ret = script_runner.run('tnefparse', '-b', 'tests/examples/one-file.tnef')
    assert ret.stderr == 'No body found\n'
    assert not ret.success


def test_cmdline_body(script_runner):
    ret = script_runner.run('tnefparse', '-b', 'tests/examples/unicode-mapi-attr.tnef')
    assert len(ret.stdout) == 12
    assert ret.stderr == ''
    assert ret.success


def test_cmdline_no_htmlbody(script_runner):
    ret = script_runner.run('tnefparse', '-hb', 'tests/examples/one-file.tnef')
    assert ret.stderr == 'No HTML body found\n'
    assert not ret.success


def test_cmdline_htmlbody(script_runner):
    ret = script_runner.run('tnefparse', '-hb', 'tests/examples/body.tnef')
    assert len(ret.stdout) == 5358
    assert ret.stderr == ''
    assert ret.success


def test_cmdline_no_rtfbody(script_runner):
    ret = script_runner.run('tnefparse', '-rb', 'tests/examples/one-file.tnef')
    assert ret.stderr == 'No RTF body found\n'
    assert not ret.success


def test_cmdline_rtfbody(script_runner):
    ret = script_runner.run('tnefparse', '-rb', 'tests/examples/rtf.tnef')
    assert len(ret.stdout) == 593
    assert ret.stderr == ''
    assert ret.success


def test_dump(script_runner):
    ret = script_runner.run('tnefparse', '-d', 'tests/examples/two-files.tnef')
    assert ret.success
    dump = json.loads(ret.stdout)
    assert sorted(list(dump.keys())) == ['attachments', 'attributes', 'extended_attributes']
    assert len(dump['attachments']) == 2


@patch.object(sys, 'argv', ['tnefparse', 'tests/examples/two-files.tnef'])
@patch.object(tnefparse.cmdline, "TNEF")
def test_handle_value_errors(mocked_TNEF):
    """Make sure e.g. invalid TNEF signatures are handled"""
    # If there is a suitable TNEF file with an e.g. invalid TNEF signature,
    # this test should be updated.
    mocked_TNEF.side_effect = ValueError("Value not allowed.")
    with pytest.raises(SystemExit):
        from tnefparse.cmdline import tnefparse
        tnefparse()
