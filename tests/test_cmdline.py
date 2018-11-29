
def test_cmdline_overview(script_runner):
    ret = script_runner.run('tnefparse', '-o', 'tests/examples/body.tnef')
    assert ret.success
    assert "Overview of tests/examples/body.tnef" in ret.stdout
    assert ret.stderr == ''

def test_cmdline_attch_extract(script_runner):
    ret = script_runner.run('tnefparse', '-a', 'tests/examples/one-file.tnef')
    assert ret.stderr == 'Successfully wrote 1 files\n'
