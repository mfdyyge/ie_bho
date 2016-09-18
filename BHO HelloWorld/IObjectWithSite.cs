using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;

namespace IE
{
    [           
       ComVisible (true),   
       InterfaceType (ComInterfaceType.InterfaceIsIUnknown),
       Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")
    ]

    //public interface IObjectWithSite
    //{       
    //      [PreserveSig]
    //      void SetSite([MarshalAs(UnmanagedType.IUnknown)]object site);       
    //}


    public interface IObjectWithSite
    {
        [PreserveSig]
        int SetSite([MarshalAs(UnmanagedType.IUnknown)]object site);
        [PreserveSig]
        int GetSite(ref Guid guid, out IntPtr ppvSite);
    }
    
}
//public int  SetSite(object site)
//{
    //if(site != null )
//    {
//       webBrowser = (WebBrowser)site;
//       webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEvenHandler(this.OnDocumentComplete);
//    }
//    else 
//    {
//       webBrowser.DocumentComplete -= new DWebBrowserEvents2_DocumentCompleteEvenHandler(this.OnDocumentComplete); 
//       webBrowser= null;
//    }
//    return 0;
//}
//public int GetSite(ref Guid guid, out IntPtr ppvSite)
//{
//  IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
//  int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
//  Marshal.Release(punk);

//  return hr;
//}