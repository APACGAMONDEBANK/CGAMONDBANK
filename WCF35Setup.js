// Configures Windows Communication Foundation 3.5 for WebMatrix
// 
// Usage: 
//    WCF35Setup.js [un]install

function WCFHandler(name, path, mode, bitness)
{
    this._name = name;
    this._path = path;
    this._mode = mode;
    this._bitness = bitness;
}

var moduleName = "ServiceModel";

var wcfHandlers = new Array( new WCFHandler("svc-Integrated", "*.svc", "integrated", null),
                             new WCFHandler("rules-Integrated", "*.rules", "integrated", null),
                             new WCFHandler("xoml-Integrated", "*.xoml", "integrated", null),

                             new WCFHandler("svc-ISAPI-2.0", "*.svc", "classic", "x86"),
                             new WCFHandler("rules-ISAPI-2.0", "*.rules", "classic", "x86"),
                             new WCFHandler("xoml-ISAPI-2.0", "*.xoml", "classic", "x86"),

                             new WCFHandler("svc-ISAPI-2.0-64", "*.svc", "classic", "x64"),
                             new WCFHandler("rules-64-ISAPI-2.0", "*.rules", "classic", "x64"),
                             new WCFHandler("xoml-64-ISAPI-2.0", "*.xoml", "classic", "x64"));

//
// main start
//
try { 
    var mode = ParseArguments();
    if (mode == "install")
    {
        UninstallWCF();
        InstallWCF(); 
        WScript.Echo("WCF 3.5 has been configured for IIS Express.");
    }
    else if (mode == "uninstall")
    {
        UninstallWCF();
        WScript.Echo("WCF 3.5 has been uninstalled from IIS Express.");
    }
    else
    {
        PrintUsage();
    }
} 
catch(e) { 
    WScript.Echo("An error occurred:\r\n " + e.description); 
} 
//
// main end
//

function InstallWCF() { 
    var adminManager = GetAdminManager(); 
 
    AddModule(adminManager);
    AddHandlers(adminManager);
 
    adminManager.CommitChanges(); 
} 

function UninstallWCF() {
    var adminManager = GetAdminManager(); 
    var moduleSection = adminManager.GetAdminSection("system.webServer/modules", "MACHINE/WEBROOT/APPHOST");

    var modulePosition = FindElement(moduleSection.Collection, "add", ["name", moduleName]); 
    if (modulePosition != -1) 
    {
      moduleSection.Collection.DeleteElement(modulePosition); 
    }

    var handlerSection = adminManager.GetAdminSection("system.webServer/handlers", "MACHINE/WEBROOT/APPHOST");

    for (i = 0; i < wcfHandlers.length; i++)
    {
        var svcPosition = FindElement(handlerSection.Collection, "add", ["name", wcfHandlers[i]._name]); 
        if (svcPosition != -1) 
        {
          handlerSection.Collection.DeleteElement(svcPosition); 
        }
    }
    
    adminManager.CommitChanges(); 
}

function AddModule(adminManager)
{
    var moduleSection = adminManager.GetAdminSection("system.webServer/modules", "MACHINE/WEBROOT/APPHOST");
    var element = moduleSection.Collection.CreateNewElement("add"); 

    element.Properties.Item("name").Value = moduleName; 
    element.Properties.Item("type").Value = "System.ServiceModel.Activation.HttpModule, System.ServiceModel, Version=3.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
    element.Properties.Item("preCondition").Value = "managedHandler,runtimeVersionv2.0";

    moduleSection.Collection.AddElement(element, -1); 
}

function AddHandlers(adminManager)
{
    var handlerSection = adminManager.GetAdminSection("system.webServer/handlers", "MACHINE/WEBROOT/APPHOST");

    for (var i = 0; i < wcfHandlers.length; i++) 
    {
        if (wcfHandlers[i]._mode == "integrated") 
        {
            AddIntegratedHandler(handlerSection, wcfHandlers[i]._name, wcfHandlers[i]._path);
        }
        else if (wcfHandlers[i]._mode == "classic") 
        {
            AddISAPIHandler(handlerSection, wcfHandlers[i]._name, wcfHandlers[i]._path, wcfHandlers[i]._bitness );
        }
        else 
        {
            throw new Error("Unrecognized mode [" + wcfHandlers[i]._mode + "]");
        }
    }
}

function AddIntegratedHandler(section, name, path)
{
    var element = section.Collection.CreateNewElement("add"); 
    element.Properties.Item("name").Value = name; 
    element.Properties.Item("path").Value = path; 
    element.Properties.Item("verb").Value = "*"; 
    element.Properties.Item("type").Value = "System.ServiceModel.Activation.HttpHandler, System.ServiceModel, Version=3.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089";
    element.Properties.Item("preCondition").Value = "integratedMode,runtimeVersionv2.0"; 
    section.Collection.AddElement(element, 0); 
}

function AddISAPIHandler(section, name, path, bitness)
{
    var element = section.Collection.CreateNewElement("add");
    var scriptProcessor = null;
    var preCondition = null;

    if (bitness == "x86") 
    {
        scriptProcessor = "%SystemRoot%\\Microsoft.NET\\Framework\\v2.0.50727\\aspnet_isapi.dll";
        preCondition = "classicMode,runtimeVersionv2.0,bitness32"; 
    }
    else if (bitness == "x64") 
    {
        scriptProcessor = "%SystemRoot%\\Microsoft.NET\\Framework64\\v2.0.50727\\aspnet_isapi.dll";
        preCondition = "classicMode,runtimeVersionv2.0,bitness64"; 
    }
    else 
    {
        throw new Error("Unrecognized bitness [" + bitness + "]");
    }

    element.Properties.Item("name").Value = name; 
    element.Properties.Item("path").Value = path; 
    element.Properties.Item("verb").Value = "*"; 
    element.Properties.Item("modules").Value = "IsapiModule";
    element.Properties.Item("scriptProcessor").Value = scriptProcessor;
    element.Properties.Item("preCondition").Value = preCondition;

    section.Collection.AddElement(element, 0); 
}

function GetAdminManager()
{
    try
    {
        var vermg = new ActiveXObject("Microsoft.IIS.VersionManager");
        var exp = vermg.GetVersionObject("10.0", 1);
        return adminManager = exp.CreateObjectFromProgId("Microsoft.ApplicationHost.WritableAdminManager");
    }
    catch(e)
    {
        throw new Error("Unable to create WritableAdminManager.\r\n Please ensure that IIS Express is installed properly.\r\n\r\n  " + e.description);
    }
}

function FindElement(collection, elementTagName, valuesToMatch) 
{ 
    for (var i = 0; i < collection.Count; i++) 
    { 
        var element = collection.Item(i); 
         
        if (element.Name == elementTagName) 
        { 
            var matches = true; 
            for (var iVal = 0; iVal < valuesToMatch.length; iVal += 2) 
            { 
                var property = element.GetPropertyByName(valuesToMatch[iVal]); 
                var value = property.Value; 
                if (value != null) 
                { 
                    value = value.toString(); 
                } 
                if (value != valuesToMatch[iVal + 1]) 
                { 
                    matches = false; 
                    break; 
                } 
            } 
            if (matches) 
            { 
                return i; 
            } 
        } 
    } 
     
    return -1; 
}

function ParseArguments()
{
    var mode = "";
    
    if (WScript.Arguments.Count() > 0)
    {
        if (WScript.Arguments.Item(0).toLowerCase() == "install")
        {
            mode="install";
        }
        else if (WScript.Arguments.Item(0).toLowerCase() == "uninstall")
        {
            mode="uninstall";
        }
    }
    
    return mode;
}

function PrintUsage()
{
    WScript.Echo("Usage:\r\n   WCF35Setup.js <cmd>\r\n\r\nDescription:\r\nAdministration utility that enables configuation of WCF 3.5 for IIS Express\r\n\r\nSupported Commands:\r\n install, uninstall\r\n\r\nSamples:\r\n WCF35Setup.js install\r\n WCF35Setup.js uninstall");
}

// SIG // Begin signature block
// SIG // MIIddAYJKoZIhvcNAQcCoIIdZTCCHWECAQExCzAJBgUr
// SIG // DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
// SIG // gjcCAR4wJAIBAQQQEODJBs441BGiowAQS9NQkAIBAAIB
// SIG // AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFIi1ZLl7QdBL
// SIG // By8iPp3D6iTn0GZwoIIYVDCCBMIwggOqoAMCAQICEzMA
// SIG // AAC7tnckcUogACAAAAAAALswDQYJKoZIhvcNAQEFBQAw
// SIG // dzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
// SIG // b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
// SIG // Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWlj
// SIG // cm9zb2Z0IFRpbWUtU3RhbXAgUENBMB4XDTE2MDkwNzE3
// SIG // NTg0N1oXDTE4MDkwNzE3NTg0N1owgbIxCzAJBgNVBAYT
// SIG // AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
// SIG // EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
// SIG // cG9yYXRpb24xDDAKBgNVBAsTA0FPQzEnMCUGA1UECxMe
// SIG // bkNpcGhlciBEU0UgRVNOOjBERTgtMkRDNS0zQ0E5MSUw
// SIG // IwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
// SIG // aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
// SIG // AQEAuPPlN/LC8UDbvDjtcJ0/Zinh7/n/48Wn0xV88fVR
// SIG // efSCv5/4y7+ByF7fLnMtyowUzOSX+g9J0JSOIUos465+
// SIG // pTutavlfKGYmGIY4EdufK4PDAyKz6WTMSujpg9HdcLZL
// SIG // B1tXBttWCoOY4InU9oZkAaYi/Qq2t3HGOZHrtJJejn2R
// SIG // aI+AxdItCUVRUstQvYCj+B/edFyI5QNmKE1jdC7hxB2D
// SIG // L8A9ZjqQKtnVcD2YesXjZ6VfFP6i53sm/SP0qmP/O7bG
// SIG // Qp+BR+0d8dE4CtnkcKVl38Bm1G9Uf1Ey3dtGkPpsuDVs
// SIG // ErJ3XhvtL7GMGheNIXgdpaNU4ZgcqzkwZhKLEwIDAQAB
// SIG // o4IBCTCCAQUwHQYDVR0OBBYEFI/cjV/uPkfBme63MuDH
// SIG // ZyR1zipLMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuz
// SIG // K7DDNbMPMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly9j
// SIG // cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
// SIG // L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYB
// SIG // BQUHAQEETDBKMEgGCCsGAQUFBzAChjxodHRwOi8vd3d3
// SIG // Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jvc29m
// SIG // dFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYB
// SIG // BQUHAwgwDQYJKoZIhvcNAQEFBQADggEBAHDCPEND8UFb
// SIG // 8WUg9X47M7g+lRf6j2YymY/Vi041o2U9pPZcxsr0i+78
// SIG // vSe5Z8jAE625uabagH3aJzcoC2gZtOTJ6yL4DyWCvPcM
// SIG // kyMktbu9a6O8/hFaxj2imYg/lrx/Elj/jlmUDX+eXxC/
// SIG // LWaPmtaB5t5VAEYua0fdejkL2+FJwhDclYPCH5rqZ8Kc
// SIG // R3JiAVFmbUygLiprl/3PBts+CifFoa7mIzw1mEFya0Ob
// SIG // ez4WWtfRnpDjrYoIHvRCiAFoWVCZMPX530ZtiBV9IFrT
// SIG // s0V1TYUddPSEmnGzcXdKFG6EWB5pasBqzs+uP90g3Tnj
// SIG // ez47y8BZk4bqFyLMa40A/DgwggYBMIID6aADAgECAhMz
// SIG // AAAAxOmJ+HqBUOn/AAAAAADEMA0GCSqGSIb3DQEBCwUA
// SIG // MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
// SIG // dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
// SIG // aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
// SIG // Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwHhcN
// SIG // MTcwODExMjAyMDI0WhcNMTgwODExMjAyMDI0WjB0MQsw
// SIG // CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
// SIG // MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
// SIG // b2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNyb3Nv
// SIG // ZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUA
// SIG // A4IBDwAwggEKAoIBAQCIirgkwwePmoB5FfwmYPxyiCz6
// SIG // 9KOXiJZGt6PLX4kvOjMuHpF4+nypH4IBtXrLGrwDykbr
// SIG // xZn3+wQd8oUK/yJuofJnPcUnGOUoH/UElEFj7OO6FYzt
// SIG // E5o13jhwVG877K1FCTBJwb6PMJkMy3bJ93OVFnfRi7uU
// SIG // xwiFIO0eqDXxccLgdABLitLckevWeP6N+q1giD29uR+u
// SIG // Ype/xYSxkK7WryvTVPs12s1xkuYe/+xxa8t/CHZ04BBR
// SIG // SNTxAMhITKMHNeVZDf18nMjmWuOF9daaDx+OpuSEF8HW
// SIG // yp8dAcf9SKcTkjOXIUgy+MIkogCyvlPKg24pW4HvOG6A
// SIG // 87vsEwvrAgMBAAGjggGAMIIBfDAfBgNVHSUEGDAWBgor
// SIG // BgEEAYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUy9Zi
// SIG // hM9gOer/Z8Jc0si7q7fDE5gwUgYDVR0RBEswSaRHMEUx
// SIG // DTALBgNVBAsTBE1PUFIxNDAyBgNVBAUTKzIzMDAxMitj
// SIG // ODA0YjVlYS00OWI0LTQyMzgtODM2Mi1kODUxZmEyMjU0
// SIG // ZmMwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1
// SIG // ApUwVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5t
// SIG // aWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljQ29kU2ln
// SIG // UENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcB
// SIG // AQRVMFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWlj
// SIG // cm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljQ29kU2ln
// SIG // UENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8E
// SIG // AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQAGFh/bV8JQyCNP
// SIG // olF41+34/c291cDx+RtW7VPIaUcF1cTL7OL8mVuVXxE4
// SIG // KMAFRRPgmnmIvGar27vrAlUjtz0jeEFtrvjxAFqUmYoc
// SIG // zAmV0JocRDCppRbHukdb9Ss0i5+PWDfDThyvIsoQzdiC
// SIG // EKk18K4iyI8kpoGL3ycc5GYdiT4u/1cDTcFug6Ay67Sz
// SIG // L1BWXQaxFYzIHWO3cwzj1nomDyqWRacygz6WPldJdyOJ
// SIG // /rEQx4rlCBVRxStaMVs5apaopIhrlihv8cSu6r1FF8xi
// SIG // ToG1VBpHjpilbcBuJ8b4Jx/I7SCpC7HxzgualOJqnWmD
// SIG // oTbXbSD+hdX/w7iXNgn+PRTBmBSpwIbM74LBq1UkQxi1
// SIG // SIV4htD50p0/GdkUieeNn2gkiGg7qceATibnCCFMY/2c
// SIG // kxVNM7VWYE/XSrk4jv8u3bFfpENryXjPsbtrj4Nsh3Kq
// SIG // 6qX7n90a1jn8ZMltPgjlfIOxrbyjunvPllakeljLEkdi
// SIG // 0iHv/DzEMQv3Lz5kpTdvYFA/t0SQT6ALi75+WPbHZ4dh
// SIG // 256YxMiMy29H4cAulO2x9rAwbexqSajplnbIvQjE/jv1
// SIG // rnM3BrJWzxnUu/WUyocc8oBqAU+2G4Fzs9NbIj86WBjf
// SIG // iO5nxEmnL9wliz1e0Ow0RJEdvJEMdoI+78TYLaEEAo5I
// SIG // +e/dAs8DojCCBgcwggPvoAMCAQICCmEWaDQAAAAAABww
// SIG // DQYJKoZIhvcNAQEFBQAwXzETMBEGCgmSJomT8ixkARkW
// SIG // A2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEt
// SIG // MCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
// SIG // dGUgQXV0aG9yaXR5MB4XDTA3MDQwMzEyNTMwOVoXDTIx
// SIG // MDQwMzEzMDMwOVowdzELMAkGA1UEBhMCVVMxEzARBgNV
// SIG // BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
// SIG // HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEh
// SIG // MB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
// SIG // MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
// SIG // n6Fssd/bSJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afR
// SIG // TK10MCAR6wfVVJUVSZQbQpKumFwwJtoAa+h7veyJBw/3
// SIG // DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36
// SIG // Rc8J0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9V
// SIG // UgzkAyz7apCQMG6H81kwnfp+1pez6CGXfvjSE/MIt1Nt
// SIG // UrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmCn4Tt
// SIG // VXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987
// SIG // g5OhYQK1HQ2x/PebsgHOIktU//kFw8IgCwIDAQABo4IB
// SIG // qzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU
// SIG // IzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGG
// SIG // MBAGCSsGAQQBgjcVAQQDAgEAMIGYBgNVHSMEgZAwgY2A
// SIG // FA6sgmBAVieX5SUT/CrhClOVWeSkoWOkYTBfMRMwEQYK
// SIG // CZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJ
// SIG // bWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9v
// SIG // dCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmCEHmtFqFKoKWt
// SIG // THNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDov
// SIG // L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
// SIG // dHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUF
// SIG // BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5t
// SIG // aWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRS
// SIG // b290Q2VydC5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgw
// SIG // DQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0rXfb
// SIG // znlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiAT
// SIG // UJndekDiV7uvWJoc4R0Bhqy7ePKL0Ow7Ae7ivo8KBciN
// SIG // SOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSokInD79vzk
// SIG // eJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiL
// SIG // ID/iBxoy+cwxSnYxPStyC8jqcD3/hQoT38IKYY7w17gX
// SIG // 606Lf8U1K16jv+u8fQtCe9RTciHuMMq7eGVcWwEXChQO
// SIG // 0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oY
// SIG // emIp76KbKTQGdxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOk
// SIG // qHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7OwTWfIn0L/gHkh
// SIG // gJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgk
// SIG // X1rqVu+m3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmh
// SIG // dmG0bqETpr+qR/ASb/2KMmyy/t9RyIwjyWa9nR2HEmQC
// SIG // PS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZ
// SIG // kocTdSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVv
// SIG // iQt9yGnI5m7qp5fOMSn/DsVbXNhNG6HY+i+ePy5VFmvJ
// SIG // E6P9MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
// SIG // hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
// SIG // BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
// SIG // HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEy
// SIG // MDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
// SIG // dGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5
// SIG // WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
// SIG // MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
// SIG // bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
// SIG // aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBTaWdu
// SIG // aW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOC
// SIG // Ag8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGf
// SIG // Qhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDI
// SIG // OdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv
// SIG // 2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13Y
// SIG // xC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT
// SIG // +OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
// SIG // 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
// SIG // kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXD
// SIG // OW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVf
// SIG // GCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4ji
// SIG // JV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bys
// SIG // AoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
// SIG // JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeB
// SIG // e+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx
// SIG // 7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+HggWCwT
// SIG // XWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEA
// SIG // AaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1Ud
// SIG // DgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEE
// SIG // AYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYw
// SIG // DwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToC
// SIG // MZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBL
// SIG // hklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
// SIG // bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
// SIG // MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEF
// SIG // BQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
// SIG // aS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNf
// SIG // MjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcu
// SIG // AzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
// SIG // b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMu
// SIG // aHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABf
// SIG // AHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBuAHQA
// SIG // LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou
// SIG // 09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+
// SIG // vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzS
// SIG // Gksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlE
// SIG // PXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6V
// SIG // oCo/KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu
// SIG // 5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
// SIG // STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
// SIG // mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38c
// SIG // bxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HN
// SIG // pZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7f
// SIG // QccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AA
// SIG // KcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
// SIG // S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8
// SIG // MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7
// SIG // qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJ
// SIG // UnMTDXpQzTGCBIwwggSIAgEBMIGVMH4xCzAJBgNVBAYT
// SIG // AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
// SIG // EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
// SIG // cG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2Rl
// SIG // IFNpZ25pbmcgUENBIDIwMTECEzMAAADE6Yn4eoFQ6f8A
// SIG // AAAAAMQwCQYFKw4DAhoFAKCBoDAZBgkqhkiG9w0BCQMx
// SIG // DAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
// SIG // KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUJrsyGzhG
// SIG // H4/RmNYQ5vepphlxYsswQAYKKwYBBAGCNwIBDDEyMDCg
// SIG // GIAWAEkASQBTACAARQB4AHAAcgBlAHMAc6EUgBJodHRw
// SIG // Oi8vd3d3Lmlpcy5uZXQwDQYJKoZIhvcNAQEBBQAEggEA
// SIG // Kndw2rFEJGL/Licyd0RD9i3H3mVQx29UVyZq/b7E+fVA
// SIG // CDleMjR8WM9iDGhxLPwl0WfHR63jnfp4SKUggV1w/NkU
// SIG // VEBs3LrhJb9jAanoOsm9iea+YoQu9Gd5FFx1y+zJYYBr
// SIG // 6IQr4gM+Y4NtkcKQkBDFSkkfEJBmtfro/J7RncLj7eO/
// SIG // qgOMRVVT0SLuLT4ZnDkOHL9covntK/E9SeQV2n9mpQuO
// SIG // 3Vkph3Vzg9/l7DD2jnIO/zJmumN9XQBSqjUFDfmgxC5V
// SIG // lu902qNs7lGh3H0V2bvE6LkVnVjAv3DAkeUQIgv3yP3t
// SIG // mK8Hz4mo89/Mi3GmH5kLb4xb03+MY83iqaGCAigwggIk
// SIG // BgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYD
// SIG // VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
// SIG // A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
// SIG // IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQg
// SIG // VGltZS1TdGFtcCBQQ0ECEzMAAAC7tnckcUogACAAAAAA
// SIG // ALswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
// SIG // hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE4MDMyODIx
// SIG // MzAwN1owIwYJKoZIhvcNAQkEMRYEFLKAEbD6eAU2psig
// SIG // qCHJPl44csjqMA0GCSqGSIb3DQEBBQUABIIBAFBFKvmM
// SIG // xSAxieWW5aiTCr2//VIVuzC1uviNEldU6FDX0q44s60N
// SIG // xO6B3ERSv9p8DaZ7ryTW6LdNSCI/E9vBwlReLuI0tzBB
// SIG // t5HLHJqctFiXpe1IYg/73XQpf3NJMKB3CYQIvXs36vG9
// SIG // yNZovaJbs41Zi4Vn/IQpOpRoR/WKV6wmcjaci70tfP23
// SIG // snc05N4fYEfhMSFyiwr5+qKLAXKDMPE5k/48hHIE71s7
// SIG // oYWpduL/DPrZwXHAiNnSQZ0VMTOn50Lh3os+J4GY8mlu
// SIG // AWO3ndw2ercylV6Mm8Vq+DrOdkGvDlgyfu6dkMYEcm33
// SIG // 9uCaBCD40n9Y/cMi3E/705+UfBc=
// SIG // End signature block
