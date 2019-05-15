﻿Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Runtime.InteropServices

Public Class QDIDataAccessInstaller

    ''' <summary>
    ''' default constructor for class QDIDataAccessInstaller
    ''' </summary>
    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub

    ''' <summary>
    ''' QDIDataAccessInstaller Install function, call Install function in superclass
    ''' </summary>
    ''' <param name="stateSaver">System.Collections.IDictionary</param>
    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        MyBase.Install(stateSaver)
        Dim regsrv As New RegistrationServices
        regsrv.RegisterAssembly(MyBase.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase)
    End Sub

    ''' <summary>
    ''' QDIDataAccessInstaller Uninstall function, call Uninstall function in superclass
    ''' </summary>
    ''' <param name="savedState">System.Collections.IDictionary</param>
    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
        MyBase.Uninstall(savedState)
        Dim regsrv As New RegistrationServices
        regsrv.UnregisterAssembly(MyBase.GetType().Assembly)
    End Sub

End Class
