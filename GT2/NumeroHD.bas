Attribute VB_Name = "NumeroHD"
'---------------------------------------------------------------------------
'Form de prueba para leer la etiqueta y el número de serie de un disco.
'                                                                (18/Feb/97)
'---------------------------------------------------------------------------
Option Explicit

'Declaración de la función, sólo está en el API de 32 bits
'
Private Declare Function GetVolumeInformation Lib "Kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                    ByVal lpVolumeNameBuffer As String, _
                                    ByVal nVolumeNameSize As Long, _
                                    lpVolumeSerialNumber As Long, _
                                    lpMaximumComponentLength As Long, _
                                    lpFileSystemFlags As Long, _
                                    ByVal lpFileSystemNameBuffer As String, _
                                    ByVal nFileSystemNameSize As Long) As Long


Public Function LeerNumeroHD(strUnidad As String) As String
    Dim lVSN As Long, n As Long, s1 As String, s2 As String
    Dim unidad As String
    Dim sTmp As String

    On Local Error Resume Next

    'Se debe especificar el directorio raiz
    unidad = Trim$(strUnidad)

    'Reservar espacio para las cadenas que se pasarán al API
    s1 = String$(255, Chr$(0))
    s2 = String$(255, Chr$(0))
    n = GetVolumeInformation(unidad, s1, Len(s1), lVSN, 0, 0, s2, Len(s2))
    's1 será la etiqueta del volumen
    'lVSN tendrá el valor del Volume Serial Number (número de serie del volumen)
    's2 el tipo de archivos: FAT, etc.

    'Convertirlo a hexadecimal para mostrarlo como en el Dir.
    sTmp = Hex$(lVSN)

    LeerNumeroHD = Trim$(sTmp)

End Function


Public Function Encriptar(strNumeroHD As String) As String
    Dim x, iNum, iASC As Integer
    Dim sLetra As String
    
    For x = 1 To Len(strNumeroHD)
        sLetra = Mid$(strNumeroHD, x, 1)
        iASC = Asc(sLetra)
        iNum = iNum + iASC
    Next x

    iNum = iNum * 1234567

    Encriptar = Hex$(iNum)
     
End Function

