import shutil
import tempfile


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
