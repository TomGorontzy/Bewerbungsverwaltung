@{
    # Ziel: praxisnahe Script-Analyse ohne dauerhafte Noise-Warnungen in Build/Release-Skripten.
    ExcludeRules = @(
        'PSPossibleIncorrectUsageOfRedirectionOperator',
        'PSAvoidUsingWriteHost',
        'PSUseBOMForUnicodeEncodedFile',
        'PSAvoidTrailingWhitespace'
    )
}
