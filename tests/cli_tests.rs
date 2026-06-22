#[test]
fn cli_tests() {
    trycmd::TestCases::new()
        .case("tests/output/*.toml")
        .case("tests/test.md")
        .case("tests/test_xlsm.md")
        .case("Readme.md");
}
