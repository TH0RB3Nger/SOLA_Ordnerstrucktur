﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Dieser Code wurde von einem Tool generiert.
'     Laufzeitversion:4.0.30319.42000
'
'     Änderungen an dieser Datei können falsches Verhalten verursachen und gehen verloren, wenn
'     der Code erneut generiert wird.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources

    'Diese Klasse wurde von der StronglyTypedResourceBuilder automatisch generiert
    '-Klasse über ein Tool wie ResGen oder Visual Studio automatisch generiert.
    'Um einen Member hinzuzufügen oder zu entfernen, bearbeiten Sie die .ResX-Datei und führen dann ResGen
    'mit der /str-Option erneut aus, oder Sie erstellen Ihr VS-Projekt neu.
    '''<summary>
    '''  Eine stark typisierte Ressourcenklasse zum Suchen von lokalisierten Zeichenfolgen usw.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0"),
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>
    Friend Class Resource1

        Private Shared resourceMan As Global.System.Resources.ResourceManager

        Private Shared resourceCulture As Global.System.Globalization.CultureInfo

        <Global.System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")>
        Friend Sub New()
            MyBase.New
        End Sub

        '''<summary>
        '''  Gibt die zwischengespeicherte ResourceManager-Instanz zurück, die von dieser Klasse verwendet wird.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>
        Friend Shared ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("WindowsApp1.Resource1", GetType(Resource1).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property

        '''<summary>
        '''  Überschreibt die CurrentUICulture-Eigenschaft des aktuellen Threads für alle
        '''  Ressourcenzuordnungen, die diese stark typisierte Ressourcenklasse verwenden.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>
        Friend Shared Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = Value
            End Set
        End Property

        '''<summary>
        '''  Sucht eine lokalisierte Ressource vom Typ System.Byte[].
        '''</summary>
        Friend Shared ReadOnly Property draußen() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("draußen", resourceCulture)
                Return CType(obj, Byte())
            End Get
        End Property

        '''<summary>
        '''  Sucht eine lokalisierte Ressource vom Typ System.Byte[].
        '''</summary>
        Friend Shared ReadOnly Property HighQuality__HQ_() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("HighQuality__HQ_", resourceCulture)
                Return CType(obj, Byte())
            End Get
        End Property

        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die s = {
        '''	id = &quot;E9E9A084-ED97-4B79-9B31-D0A81742FB90&quot;,
        '''	internalName = &quot;SOLA19_Kids_HighQuality (HQ)&quot;,
        '''	title = &quot;SOLA21_Teens_HighQuality (HQ)&quot;,
        '''	type = &quot;Export&quot;,
        '''	value = {
        '''		DNG_compatibilityV2 = 184680448,
        '''		collisionHandling = &quot;ask&quot;,
        '''		embeddedMetadataOption = &quot;all&quot;,
        '''		exportServiceProvider = &quot;com.adobe.ag.export.file&quot;,
        '''		exportServiceProviderTitle = &quot;Festplatte&quot;,
        '''		export_colorSpace = &quot;sRGB&quot;,
        '''		export_destinationPathPrefix = &quot;/Volumes/SSD_11_e/2020-05_Pfijuko20/Selected by Lars&quot;,
        '''		export_destinationPathS [Rest der Zeichenfolge wurde abgeschnitten]&quot;; ähnelt.
        '''</summary>
        Friend Shared ReadOnly Property HighQuality__HQ_Txt() As String
            Get
                Return ResourceManager.GetString("HighQuality__HQ_Txt", resourceCulture)
            End Get
        End Property

        '''<summary>
        '''  Sucht eine lokalisierte Ressource vom Typ System.Byte[].
        '''</summary>
        Friend Shared ReadOnly Property LowQuality__LQ_() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("LowQuality__LQ_", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Ressource vom Typ System.Byte[].
        '''</summary>
        Friend Shared ReadOnly Property RAW() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("RAW", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Ressource vom Typ System.Byte[].
        '''</summary>
        Friend Shared ReadOnly Property Veranstaltungszelt() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("Veranstaltungszelt", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
    End Class
End Namespace
