Option Explicit On
Option Strict On
''' <summary>
''' An interface for use with getting the ICursor into the QdiRecord
''' </summary>
Public Interface ICursorToQdiRecord
    ReadOnly Property RecordList() As List(Of Qdi.BusinessLogic.IQdiRecord)
    Function ConvertFeatureIntoQdiRecord(ByVal pFeature As ESRI.ArcGIS.Geodatabase.IFeature) As Qdi.BusinessLogic.IQdiRecord
End Interface
