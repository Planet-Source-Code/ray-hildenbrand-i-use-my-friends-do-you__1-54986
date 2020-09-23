VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPublicClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This example illustrates how and why a Friend
'method is used in vb6 to pass information between a form to
'as class without exposing another property in the component's
'object model. While this a limited example of using Friend methods (properties,etc...)
'it should express the point.... A friend is used to pass information between modules in
'a vb project while keeping it's scope to that of the component and not the caller (project)
'so to summarize, a friend is public to the component/project it resides in and is private (or hidden)
'and is not exposed to the project that references the component.

'In this example I chose to only expose a property Get for the class
'to illustrate a situation where you would only expect and permit the
'property value to be changed by a built in UI (and/or code) and expose
'only the current value to the developer of the application logic who uses the compiled component.

'all and all, a very simple example but one that might promote proper habits for the
'component builders out there.

'any questions can be forwarded to rayhildenbrand@vibrantinnovations.com
'http://www.vibrantinnovations.com/vlt.asp    -Vibrant Layout Tools
'Copyright 2004 Ray Hildenbrand and VibrantInnovations

''No votes needed, just build some cool components..................

'private global to hold the information passed between our friend
Private mCurrentSelectionString As String
Public Event SelectionChanged(newValue As String) 'event just to help show it in action


Public Function GetPropertyFromUI() As String
    Dim tmpStringSelection As String
    Dim fEx As New frmExample ' instantiate the form that will set the friend
    
    tmpStringSelection = mCurrentSelectionString 'set tmp var in this example to check if we need to raiseevent
    
    fEx.Initialize Me 'set up the link between the built in form and our class
    fEx.Show vbModal 'show the ui to get the results
    
    ''//--at this point when the form returns
    ''//--the private variable mCurrentSelectionString would be
    ''//--populated with any changes (if any) made in the UI form
    If Not tmpStringSelection = mCurrentSelectionString Then
        RaiseEvent SelectionChanged(mCurrentSelectionString)
    Else
        'whatever
    End If
    
    Set fEx = Nothing
End Function

'//----Here is the friend method. Caller projects cannot see this procedure.
'//----Objects (controls,classes, modules, forms, etc...) within the scope
'//----of this object can however see and call this method with no problem
Friend Sub DefineSelection(newSelection As String)
     mCurrentSelectionString = newSelection
End Sub

'//----In this example I am only exposing a property get for
'//----a property that is read only unless it gets set through the friend
'//----method within this project ().
'//----You may have a different reason to do this, think outside the code example

'//----So projects that reference this project as a compiled dll will only see
'//----this method,our UI showing sub and the test event, but we have an effective way
'//----to pass information between our object while maintaining a nice tight object model.


Public Property Get CurrentSelection() As String
    CurrentSelection = mCurrentSelectionString
End Property

