
@set command=cscript //nologo TestLauncher.vbs

:: configure doskey macros

@doskey test=%command%
@doskey wip=%command% Chooser.spec.wip.vbs

@cmd /k %command%
