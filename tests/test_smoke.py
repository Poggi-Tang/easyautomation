# tests/test_smoke.py
def test_smoke():
    try:
        import easy_uiauto
        from easy_uiauto import record
        record.record_help()
        assert True
    except:
        assert False
