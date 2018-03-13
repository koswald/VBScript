Option Explicit : Initialize
With New TestingFramework
    .describe "IdleTimer.dll/IdleTimer class"
        Dim it : Set it = CreateObject("VBScripting.IdleTimer")
    .it "should initialize the default reset period"
        .AssertEqual it.ResetPeriod, 30000
    .it "should initialize the desired state"
        .AssertEqual it.DesiredState, &h80000000
    .it "should initialize DisplayRequired"
        .AssertEqual it.DisplayRequired, False
    .it "should initialize SystemRequired"
        .AssertEqual it.SystemRequired, False
    .it "should return a double for ResetPeriod"
        .AssertEqual TypeName(it.ResetPeriod), "Double"
    .it "should return a double for CurrentState"
        .AssertEqual TypeName(it.CurrentState), "Double"
    .it "should return a double for DesiredState"
        .AssertEqual TypeName(it.DesiredState), "Double"
    'set DisplayRequired first to test SystemRequired.set logic
    .it "should set the SystemRequired desired state,  !SysReq'd, !DispReq'd"
        it.DisplayRequired = False
        it.SystemRequired = False
        .AssertEqual it.DesiredState, &h80000000
    .it "should set the SystemRequired desired state,  !SysReq'd,  DispReq'd"
        it.DisplayRequired = True
        it.SystemRequired = False
        .AssertEqual it.DesiredState, &h80000002
    .it "should set the SystemRequired desired state,   SysReq'd,  DispReq'd"
        it.DisplayRequired = True
        it.SystemRequired = True
        .AssertEqual it.DesiredState, &h80000003
    .it "should set the SystemRequired desired state,   SysReq'd, !DispReq'd"
        it.DisplayRequired = False
        it.SystemRequired = True
        .AssertEqual it.DesiredState, &h80000001
    'set SystemRequired first to test DisplayRequired.set logic
    .it "should set the DisplayRequired desired state, !SysReq'd, !DispReq'd"
        it.SystemRequired = False
        it.DisplayRequired = False
        .AssertEqual it.DesiredState, &h80000000
    .it "should set the DisplayRequired desired state, !SysReq'd,  DispReq'd"
        it.SystemRequired = False
        it.DisplayRequired = True
        .AssertEqual it.DesiredState, &h80000002
    .it "should set the DisplayRequired desired state,  SysReq'd,  DispReq'd"
        it.SystemRequired = True
        it.DisplayRequired = True
        .AssertEqual it.DesiredState, &h80000003
    .it "should set the DisplayRequired desired state,  SysReq'd, !DispReq'd"
        it.SystemRequired = True
        it.DisplayRequired = False
        .AssertEqual it.DesiredState, &h80000001
    'test DisplayRequired.get
    .it "should isolate the DisplayRequired flag (1)"
        it.DesiredState = &h80000001
        .AssertEqual it.DisplayRequired, False
     .it "should isolate the DisplayRequired flag (2)"
        it.DesiredState = &h80000002
        .AssertEqual it.DisplayRequired, True
     .it "should isolate the DisplayRequired flag (3)"
        it.DesiredState = &h80000004
        .AssertEqual it.DisplayRequired, False
    'test SystemRequired.get
    .it "should isolate the SystemRequired flag (1)"
        it.DesiredState = &h80000000
        .AssertEqual it.SystemRequired, False
     .it "should isolate the SystemRequired flag (2)"
        it.DesiredState = &h80000001
        .AssertEqual it.SystemRequired, True
     .it "should isolate the SystemRequired flag (3)"
        it.DesiredState = &h80000002
        .AssertEqual it.SystemRequired, False
End With

it.Dispose
Set it = Nothing

Sub Initialize
    With CreateObject("VBScripting.Includer")
        ExecuteGlobal .read("TestingFramework")
    End With
End Sub
