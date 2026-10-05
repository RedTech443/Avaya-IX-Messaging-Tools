@{
    # IXM-Tools is an interactive Windows PowerShell 5.1 console application.
    # These exclusions are intentional UI/naming conventions only.
    # Correctness/security rules remain enabled.
    Severity = @('Error', 'Warning')

    ExcludeRules = @(
        'PSAvoidUsingWriteHost'
        'PSUseSingularNouns'
        'PSUseApprovedVerbs'
        'PSUseShouldProcessForStateChangingFunctions'
    )
}