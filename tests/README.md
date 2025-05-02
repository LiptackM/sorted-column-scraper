# Unit tests and functional tests for sorted-column-scraper


notes:
    - requires a test_file & test_corrupt_file (see fixtures) in the same folder
            as this test.  Note if you provide these, also change expected values below
    - test_corrupt-file is jsut a simple text file renamed to a xlsx so its corrupt
    - pytest.ini needs to know the root where these are,
        typically by adding "pythonpath = ."

to run with coverage (coverage suggested, initally this was 87%, when
        only skipping simple I/O in main()):
    > ppytest -rs --cov=excel_column_sorted --cov-report=term-missing 
