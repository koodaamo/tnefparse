import io
import zipfile
import json
import logging
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

import tnefparse.cmdline

EXAMPLES = (Path(__file__).resolve().parent / "examples").relative_to(Path('.').resolve())


def test_cmdline_overview(script_runner):
    ret = script_runner.run('tnefparse', '-o', str(EXAMPLES / 'body.tnef'))
    assert ret.success
    assert "Overview of tests/examples/body.tnef" in ret.stdout
    assert ret.stderr == ''
    assert ret.success


def test_cmdline_attch_extract(script_runner, tmp_path):
    ret = script_runner.run('tnefparse', '-a', '-p', str(tmp_path), str(EXAMPLES / 'one-file.tnef'))
    assert ret.stderr == 'Successfully wrote 1 files\n'
    assert ret.success


def test_cmdline_zip_extract(script_runner, tmp_path):
    ret = script_runner.run('tnefparse', '-z', '-p', str(tmp_path), str(EXAMPLES / 'duplicate_filename.tnef'))
    assert (tmp_path / "attachments.zip").is_file()
    assert ret.stderr == 'Successfully wrote attachments.zip\n'
    zip_stream = io.BytesIO((tmp_path / "attachments.zip").read_bytes())
    zip_file = zipfile.ZipFile(zip_stream)
    assert zip_file.namelist() == ['file_abcdefgh.txt', 'file_abcdefgh-2.txt', 'VIA_Nytt_14021.htm']
    assert ret.success


def test_cmdline_no_body(script_runner):
    ret = script_runner.run('tnefparse', '-b', str(EXAMPLES / 'one-file.tnef'))
    assert ret.stderr == 'No body found\n'
    assert not ret.success


def test_cmdline_body(script_runner):
    ret = script_runner.run('tnefparse', '-b', str(EXAMPLES / 'unicode-mapi-attr.tnef'))
    assert len(ret.stdout) == 12
    assert ret.stderr == ''
    assert ret.success


def test_cmdline_no_htmlbody(script_runner):
    ret = script_runner.run('tnefparse', '-hb', str(EXAMPLES / 'one-file.tnef'))
    assert ret.stderr == 'No HTML body found\n'
    assert not ret.success


def test_cmdline_htmlbody(script_runner):
    ret = script_runner.run('tnefparse', '-hb', str(EXAMPLES / 'body.tnef'))
    assert len(ret.stdout) == 5358
    assert ret.stderr == ''
    assert ret.success


def test_cmdline_no_rtfbody(script_runner):
    ret = script_runner.run('tnefparse', '-rb', str(EXAMPLES / 'one-file.tnef'))
    assert ret.stderr == 'No RTF body found\n'
    assert not ret.success


def test_cmdline_rtfbody(script_runner):
    ret = script_runner.run('tnefparse', '-rb', str(EXAMPLES / 'rtf.tnef'))
    assert len(ret.stdout) == 593
    assert ret.stderr == ''
    assert ret.success


def test_dump(script_runner):
    ret = script_runner.run('tnefparse', '-d', str(EXAMPLES / 'two-files.tnef'))
    assert ret.success
    dump = json.loads(ret.stdout)
    assert sorted(dump.keys()) == ['attachments', 'attributes', 'extended_attributes']
    assert len(dump['attachments']) == 2


@patch.object(sys, 'argv', ['tnefparse', str(EXAMPLES / 'two-files.tnef')])
@patch.object(tnefparse.cmdline, "TNEF")
def test_handle_value_errors(mocked_TNEF):
    """Make sure e.g. invalid TNEF signatures are handled"""
    # If there is a suitable TNEF file with an e.g. invalid TNEF signature,
    # this test should be updated.
    mocked_TNEF.side_effect = ValueError("Value not allowed.")
    with pytest.raises(SystemExit):
        from tnefparse.cmdline import tnefparse
        tnefparse()


def test_help_is_printed(script_runner):
    """calling `tnefparse` with no argument prints help text"""
    ret = script_runner.run('tnefparse')
    help_text = "Extract TNEF file contents. Show this help message if no arguments are given."
    assert help_text in ret.stdout
    assert "positional arguments" in ret.stdout
    if sys.version_info < (3, 10):
        assert "optional arguments" in ret.stdout
    else:
        assert "Extract TNEF file contents" in ret.stdout
        assert "-d, --dump" in ret.stdout


def test_print_overview(script_runner):
    ret = script_runner.run('tnefparse', '-o', str(EXAMPLES / 'one-file.tnef'))
    assert "Overview of tests/examples/one-file.tnef" in ret.stdout
    assert "Attachments" in ret.stdout
    assert "Objects" in ret.stdout
    assert "Properties" in ret.stdout


def test_cmdline_logging_info(caplog, capsys):
    """Logging level set to INFO returns some INFO log messages"""
    retv = tnefparse.cmdline.tnefparse(('-l', 'INFO', '-rb', str(EXAMPLES / 'rtf.tnef')))
    assert not retv
    stdout, _ = capsys.readouterr()
    assert len(stdout) == 593
    assert caplog.record_tuples == [
        ('tnefparse', logging.INFO, 'Skipping checksum for performance'),
    ]


def test_cmdline_logging_warn(caplog, capsys):
    """Logging level set to WARN returns no INFO log messages"""
    retv = tnefparse.cmdline.tnefparse(('-l', 'WARN', '-rb', str(EXAMPLES / 'rtf.tnef')))
    assert not retv
    stdout, _ = capsys.readouterr()
    assert len(stdout) == 593
    assert caplog.record_tuples == []


def test_cmdline_logging_illegal():
    """Logging level set to ILLEGAL is illegal"""
    with pytest.raises(SystemExit):
        tnefparse.cmdline.tnefparse(('-l', 'ILLEGAL', '-rb', str(EXAMPLES / 'rtf.tnef')))
