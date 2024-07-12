#[test]
fn cli_tests() {
    trycmd::TestCases::new()
        .case("tests/output/*.toml")
        .case("tests/test.md")
        .case("Readme.md");
}
