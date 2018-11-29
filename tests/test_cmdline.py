
def test_foo_bar(script_runner):
    ret = script_runner.run('tnefparse', '-o', 'tests/examples/body.tnef')
    assert ret.success
    assert "Overview of tests/examples/body.tnef" in ret.stdout
    assert ret.stderr == ''

