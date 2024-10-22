VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public position As Scripting.Dictionary

Private Sub Class_Initialize()
    Set position = New Scripting.Dictionary
    
    ' Initialize APAB dictionary
    Dim apabDict As Scripting.Dictionary
    Set apabDict = New Scripting.Dictionary
    Dim apabs As Collection
    Set apabs = New Collection
    apabs.Add "f2, up*13, tab, delete*59"
    apabs.Add "f2, up*12, tab, delete*59"
    apabs.Add "f2, up*11, tab, delete*59"
    apabs.Add "f2, up*10, tab, delete*59"
    apabs.Add "f2, up*9, tab, delete*59"
    apabs.Add "f2, up*8, tab, delete*59"
    apabs.Add "f2, up*7, tab, delete*59"
    apabs.Add "f2, up*6, tab, delete*59"
    apabs.Add "f2, up*5, tab, delete*59"
    apabs.Add "f2, up*4, tab, delete*59"
    
    With apabDict
        .Add "transaction", "f2, delete*20, apab, f12"
        .Add "plant", "home, tab, delete*2"
        .Add "partNo", "home, tab*2, delete*22"
        .Add "operationNo", "home, tab, down*3, delete*4"
        .Add "designation", apabs
    End With
    
    position.Add "APAB", apabDict
    
    ' Initialize APAG dictionary
    Dim apagDict As Scripting.Dictionary
    Set apagDict = New Scripting.Dictionary
    
    With apagDict
        .Add "transaction", "f2, delete*20, apag, f12"
        .Add "plant", "home, tab, delete*2"
        .Add "partNo", "home, tab*2, delete*22"
        .Add "operationNo", "home, tab*4, delete*4"
        .Add "sequenceNo", "home, tab*4, down, delete*4"
        .Add "opTextCode", "home, tab*4, down*2, delete*4"
        .Add "costCentre", "home, up*6, tab, delete*6"
        .Add "workCentre", "home, up*6, tab*2, delete*6"
        .Add "pl", "home, up*8, tab*2, delete*2"
        .Add "warehouse", "home, up*8, tab*3, delete*6"
        .Add "backflush", "home, up*8, tab, delete"
    End With
    
    position.Add "APAG", apagDict
    
    ' Initialize APAR dictionary
    Dim aparDict As Scripting.Dictionary
    Set aparDict = New Scripting.Dictionary
    
    With aparDict
        .Add "transaction", "f2, delete*20, apar, f12"
        .Add "plant", "home, tab, delete*2"
        .Add "partNo", "home, tab*2, delete*22"
        .Add "operationNo", "home, tab, down*3, delete*4"
        .Add "qualification", "home, tab, down*4, tab*2, delete*8"
        .Add "peGroup", "home, up*8, tab, delete*8"
    End With
    
    position.Add "APAR", aparDict
    
    ' Initialize APAS dictionary
    Dim apasDict As Scripting.Dictionary
    Set apasDict = New Scripting.Dictionary
    
    With apasDict
        .Add "transaction", "f2, delete*20, apas, f12"
        .Add "plant", "home, tab, delete*2"
        .Add "partNo", "home, tab*2, delete*22"
        .Add "operationNo", "home, tab*4, delete*4"
        .Add "bomItem", "home, tab*5, delete*4"
        .Add "component", "home, tab*7, delete*22"
        .Add "amount", "home, tab*8, delete*13"
    End With
    
    position.Add "APAS", apasDict
    
    ' Initialize APAZ dictionary
    Dim apazDict As Scripting.Dictionary
    Set apazDict = New Scripting.Dictionary
    
    With apazDict
        .Add "transaction", "f2, delete*20, apaz, f12"
        .Add "plant", "home, tab, delete*2"
        .Add "partNo", "home, tab*2, delete*22"
        .Add "operationNo", "home, tab, down*3, delete*4"
        .Add "time", "home, tab, down*6, tab*2, delete*9"
    End With
    
    position.Add "APAZ", apazDict
    
    ' Initialize APSG dictionary
    Dim apsgDict As Scripting.Dictionary
    Set apsgDict = New Scripting.Dictionary
    
    With apsgDict
        .Add "transaction", "f2, delete*20, apsg, f12"
        .Add "plant", "home, tab, delete*2"
        .Add "partNo", "home, tab*2, delete*22"
        .Add "operationNo", "home, tab, down*3, delete*4"
        .Add "sequenceNo", "home, tab, down*3, tab*4, delete*4"
        .Add "opTextCode", "home, tab, down*4, delete*4"
        .Add "costCentre", "home, tab, down*5, delete*4"
        .Add "peGroup", "home, tab, down*7, tab*2, delete*8"
        .Add "qualification", "home, up*9, tab, delete*8"
        .Add "time", "home, up*6, tab*2, delete*9"
        .Add "pl", "home, up*2, tab"
        .Add "warehouse", "home, up*2, tab*2"
    End With
    
    position.Add "APSG", apsgDict
    
    ' Initialize MAGD dictionary
    Dim magdDict As Scripting.Dictionary
    Set magdDict = New Scripting.Dictionary
    
    With magdDict
        .Add "transaction", "f2, delete*20, magd, f12"
        .Add "partNo", "home, tab*2, delete*22"
        .Add "projectNumber", "home, tab*2, down*2, delete*35"
        .Add "defcod", "home, tab*2, down*2, tab, delete*6"
    End With
    
    position.Add "MAGD", magdDict
End Sub
