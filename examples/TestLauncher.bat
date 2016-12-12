
@set command=cscript //nologo TestLauncher.vbs

@doskey test=%command%
@doskey test2=%command% Chooser.spec.wip.vbs

@cmd /k %command%
