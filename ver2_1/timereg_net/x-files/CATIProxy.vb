﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:2.0.50727.4927
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'This source code was auto-generated by wsdl, Version=2.0.50727.42.
'
Namespace CATIService
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="CATISoap", [Namespace]:="http://tempuri.org/")>  _
    Partial Public Class CATI
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private upDateCatiTimerOperationCompleted As System.Threading.SendOrPostCallback
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            'Me.Url = "http://localhost/timeout_xp/timereg_net/catiws.asmx"
            Me.Url = "https://outzource.dk/timeout_xp/wwwroot/ver2_1/timereg_net/catiws.asmx"
        End Sub
        
        '''<remarks/>
        Public Event upDateCatiTimerCompleted As upDateCatiTimerCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/upDateCatiTimer", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function upDateCatiTimer(ByVal ds As System.Data.DataSet) As <System.Xml.Serialization.XmlArrayItemAttribute("DataSet")> System.Data.DataSet()
            Dim results() As Object = Me.Invoke("upDateCatiTimer", New Object() {ds})
            Return CType(results(0),System.Data.DataSet())
        End Function
        
        '''<remarks/>
        Public Function BeginupDateCatiTimer(ByVal ds As System.Data.DataSet, ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("upDateCatiTimer", New Object() {ds}, callback, asyncState)
        End Function
        
        '''<remarks/>
        Public Function EndupDateCatiTimer(ByVal asyncResult As System.IAsyncResult) As System.Data.DataSet()
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),System.Data.DataSet())
        End Function
        
        '''<remarks/>
        Public Overloads Sub upDateCatiTimerAsync(ByVal ds As System.Data.DataSet)
            Me.upDateCatiTimerAsync(ds, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub upDateCatiTimerAsync(ByVal ds As System.Data.DataSet, ByVal userState As Object)
            If (Me.upDateCatiTimerOperationCompleted Is Nothing) Then
                Me.upDateCatiTimerOperationCompleted = AddressOf Me.OnupDateCatiTimerOperationCompleted
            End If
            Me.InvokeAsync("upDateCatiTimer", New Object() {ds}, Me.upDateCatiTimerOperationCompleted, userState)
        End Sub
        
        Private Sub OnupDateCatiTimerOperationCompleted(ByVal arg As Object)
            If (Not (Me.upDateCatiTimerCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent upDateCatiTimerCompleted(Me, New upDateCatiTimerCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42")>  _
    Public Delegate Sub upDateCatiTimerCompletedEventHandler(ByVal sender As Object, ByVal e As upDateCatiTimerCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("wsdl", "2.0.50727.42"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class upDateCatiTimerCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As System.Data.DataSet()
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),System.Data.DataSet())
            End Get
        End Property
    End Class
End Namespace
