Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Public Class NotifyPropertyChanged
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub NotifyPropertyChanged(Of T)(ByRef field As T, value As T, <CallerMemberName> Optional Name As String = "")
        If Not EqualityComparer(Of T).[Default].Equals(field, value) Then
            field = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(Name))
        End If
    End Sub

    Public Sub Update()
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(String.Empty))
    End Sub

End Class