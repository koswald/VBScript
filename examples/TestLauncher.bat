
@set command=cscript //nologo TestLauncher.vbs

:: configure doskey macros

@doskey test=%command%
@doskey wip=%command% Chooser.spec.wip.vbs
@doskey wip2=%command% VBSClipboard.spec.vbs 2

@cmd /k %command%
