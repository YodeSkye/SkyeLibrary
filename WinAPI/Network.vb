
Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' DECLARATIONS
        Public Enum WLAN_INTF_OPCODE
            wlan_intf_opcode_current_connection = 7
        End Enum
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure WLAN_INTERFACE_INFO_LIST
            Public dwNumberOfItems As UInteger
            Public dwIndex As UInteger
        End Structure
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure WLAN_INTERFACE_INFO
            Public InterfaceGuid As Guid
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
            Public strInterfaceDescription As String
            Public isState As Integer
        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure DOT11_SSID
            Public uSSIDLength As UInteger
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
            Public ucSSID As Byte()
        End Structure
        <StructLayout(LayoutKind.Sequential)>
        Public Structure WLAN_ASSOCIATION_ATTRIBUTES
            Public dot11Ssid As DOT11_SSID
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=6)>
            Public dot11Bssid As Byte()
            Public dot11PhyType As Integer
            Public uDot11PhyIndex As UInteger
            Public wlanSignalQuality As UInteger
            Public ulRxRate As UInteger
            Public ulTxRate As UInteger
        End Structure
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure WLAN_CONNECTION_ATTRIBUTES
            Public isState As Integer
            Public wlanConnectionMode As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
            Public strProfileName As String
            Public wlanAssociationAttributes As WLAN_ASSOCIATION_ATTRIBUTES
            Public wlanSecurityAttributes As Integer
        End Structure

        ' API FUNCTIONS
        <DllImport("wlanapi.dll")>
        Public Shared Function WlanOpenHandle(ByVal dwClientVersion As UInteger,
                                              ByVal pReserved As IntPtr,
                                              ByRef pdwNegotiatedVersion As UInteger,
                                              ByRef phClientHandle As IntPtr) As Integer
        End Function
        <DllImport("wlanapi.dll")>
        Public Shared Function WlanEnumInterfaces(ByVal hClientHandle As IntPtr,
                                                  ByVal pReserved As IntPtr,
                                                  ByRef ppInterfaceList As IntPtr) As Integer
        End Function
        <DllImport("wlanapi.dll")>
        Public Shared Function WlanQueryInterface(ByVal hClientHandle As IntPtr,
                                                  ByVal pInterfaceGuid As Guid,
                                                  ByVal OpCode As WLAN_INTF_OPCODE,
                                                  ByVal pReserved As IntPtr,
                                                  ByRef pdwDataSize As UInteger,
                                                  ByRef ppData As IntPtr,
                                                  ByVal pWlanOpcodeValueType As UInteger) As Integer
        End Function
        <DllImport("wlanapi.dll")>
        Public Shared Sub WlanFreeMemory(ByVal pMemory As IntPtr)
        End Sub
        <DllImport("wlanapi.dll")>
        Public Shared Function WlanCloseHandle(ByVal hClientHandle As IntPtr,
                                               ByVal pReserved As IntPtr) As Integer
        End Function

    End Class

End Namespace
#Enable Warning CA1401
