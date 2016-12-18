
:: prepare to manually launch a series of tests from a console window

:: define variables

@set test=cscript //nologo TestLauncher.vbs
@set wip=%test% VBSClipboard.spec.vbs
@set wip100=%wip% 100 

@set macro1=test=%test%
@set macro2=wip=%wip%
@set macro3=wip100=%wip100%

:: define doskey macros

@doskey %macro1%
@doskey %macro2%
@doskey %macro3%

:: show macro definitions

@echo %macro1%
@echo %macro2%
@echo %macro3%
@echo.

:: start the first test

@cmd /k %wip%
