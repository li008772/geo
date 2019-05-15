Option Explicit On
Option Strict On

''' <summary>
'''     An interface (with one function signature) used by QdiRecordStandardizer.vb
''' </summary>
Public Interface IQdiRecordStandardizer
    Function Standardize(ByVal pQdirecord As Qdi.BusinessLogic.IQdiRecord, ByVal p As IDataAccess) As Qdi.BusinessLogic.IQdiRecord
End Interface
