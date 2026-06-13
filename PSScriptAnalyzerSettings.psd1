@{
    # Ziel: bekannte False-Positive-Meldung bei Native-CLI-Argumenten in release.ps1 entschärfen.
    ExcludeRules = @(
        'PSPossibleIncorrectUsageOfRedirectionOperator'
    )
}
