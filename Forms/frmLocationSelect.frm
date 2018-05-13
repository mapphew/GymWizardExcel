VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLocationSelect 
   Caption         =   "Select Your Location"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "frmLocationSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLocationSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sStoreID As String

Private m_objHTMLDoc As MSHTML.HTMLDocument

Private m_objOptions As IHTMLElementCollection



Public Property Get HTMLDoc() As MSHTML.HTMLDocument: Set HTMLDoc = m_objHTMLDoc: End Property

Public Property Get Options() As IHTMLElementCollection: Set Options = m_objOptions: End Property

Public Property Set HTMLDoc(ByVal objNewValue As HTMLDocument): Set m_objHTMLDoc = objNewValue: End Property

Public Property Set Options(ByVal objNewValue As IHTMLElementCollection): Set m_objOptions = objNewValue: End Property



Private Sub PopulateList(ByVal oHTMLDoc As HTMLDocument)
    Set m_objHTMLDoc = oHTMLDoc
    
    Dim oSelect As HTMLSelectElement
    Dim oOptionItr As HTMLOptionElement
    
    Set oSelect = m_objHTMLDoc.getElementById("stores")
    Set m_objOptions = oSelect.getElementsByTagName("option")
    
    For Each oOptionItr In m_objOptions
        Me.lBoxLocations.AddItem (oOptionItr.innerText)
    Next
    
    
    Me.Show
End Sub

Public Property Get StoreID() As String

    StoreID = m_sStoreID

End Property

Public Sub SelectLocation(ByVal oHTMLDoc As HTMLDocument)
    
    PopulateList oHTMLDoc
    
    
End Sub

Private Sub btnLogIn_Click()
    Dim objOptionItr As HTMLOptionElement
    
    For Each objOptionItr In m_objOptions
        If objOptionItr.innerText = Me.lBoxLocations.Value Then
            m_sStoreID = objOptionItr.getAttribute("value")
            Me.Hide
            Exit Sub
        End If
    Next
    
        
End Sub

Private Sub UserForm_Initialize()
    Me.Hide
End Sub
