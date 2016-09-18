using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SHDocVw;
using mshtml;

using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace IE
{
    [
       ComVisible (true),
       Guid("8a194578-81ea-4850-9911-13ba2d71efbd"),
       ClassInterface(ClassInterfaceType.None)
    ]

    public class IObjectWithSiteImp: IObjectWithSite
    {

        //调试信息开关
        protected string Debug_config ="off";
        protected string msg = "";
        protected string methodName="";

        /// <summary>
        /// 构造函数
        /// </summary>
        public IObjectWithSiteImp()
        {

        }




        #region  全局变量成员
        /********************************************************************************************************自定义实现需要的字段》时间：2016年9月14日15:46:07*/
        
        WebBrowser webBrowser;
        HTMLDocument document;

        protected SHDocVw.IWebBrowser2 iwebBrowser2; //浏览器对象
        protected DWebBrowserEvents2_Event browserEvents;//浏览器事件


        List<IHTMLElement>  iframe_arrylist = new List<IHTMLElement>();//重要*   保存找到的IFrame对象
        List<string>        iframe_name_arrylist = new List<string>();//重要*   保存找到的IFrame对象的name 属性
        Util               util=new Util() ;
        /*********************************************************************************************************/
        #endregion












	

        public void OnDucumentComplete(object pDisp, ref object URL)
        {

            this.Debug_config = util.getFilename("debug_config");/*初始化：读取Debug_Msg 配置是否==‘开’*/
            /**日志**/
            this.methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;//程序执行位置代码块的方法名
            this.msg = methodName + "进入此方法！";
            /***/
            util.log_to(this.Debug_config, this.methodName, this.msg);
            /**日志**/

            try
            {
                document = (HTMLDocument)webBrowser.Document;
                foreach (IHTMLInputElement tempElement in document.getElementsByTagName("INPUT"))
                {


                    if (((IHTMLElement)tempElement).id != null && ((IHTMLElement)tempElement).id=="su")
                    {
                        
                        /**日志**/
                        this.msg = methodName + "　baidu_id=" + ((IHTMLElement)tempElement).id;
                        util.log_to(this.Debug_config, this.methodName, this.msg);
                        /**日志**/

                        ((IHTMLElement)tempElement).setAttribute("value", "我叫周露 -"+this.Debug_config);
                    }

                    //System.Windows.Forms.MessageBox.Show(tempElement.name != null ? tempElement.name : "it sucks, no name, try id=" + ((IHTMLElement)tempElement).id  );
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }

        public void OnBeforeNavigate2(object pDisp, ref object URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancle)
        {
            document = (HTMLDocument)webBrowser.Document;
            foreach (IHTMLInputElement tempElement in document.getElementsByTagName("INPUT"))
            {
                if (tempElement.type.ToLower() == "password")
                {
                    System.Windows.Forms.MessageBox.Show(tempElement.value);
                }
            }
        }
























        /*插入远程JS填屏插件*****************************************************************************************************************************/
        /// <summary>
        /// 时间：2016年9月14日13:54:38
        /// 2：在当前网页所有的IFrame中查找（iframe.src like url ） 
        ///    a.先找到IFrame.doc
        ///    b.在DOC中插入远程JS插件
        /// </summary>
        /// <param name="like_url"></param>
        public void OnDocumentComplete_query_iframe_like_url(string like_url)
        {
            /**日志**/
            this.methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;//程序执行位置代码块的方法名
            this.msg = methodName + "进入此方法！";
            /***/
            util.log_to(this.Debug_config, this.methodName, this.msg);
            /**日志**/


            mshtml.HTMLDocument document = (mshtml.HTMLDocument)webBrowser.Document;
            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)webBrowser.Document;
            mshtml.IHTMLDocument2 doc2 = (mshtml.IHTMLDocument2)webBrowser.Document;
            mshtml.IHTMLDocument3 doc3 = (mshtml.IHTMLDocument3)webBrowser.Document;
            mshtml.IHTMLDocument4 doc4 = (mshtml.IHTMLDocument4)webBrowser.Document;
            mshtml.IHTMLDocument5 doc5 = (mshtml.IHTMLDocument5)webBrowser.Document;


            mshtml.IHTMLElement img = (mshtml.IHTMLElement)doc2.all.item("button", 0);
            int frames = doc.parentWindow.frames.length;
            int doc3_iframe_lenght = doc3.getElementsByTagName("iframe").length; //iframe 数量   //int doc2_iframe_lenght=doc2.getElementsByTagName("iframe").length; //iframe 数量 e不能正常出数
            int inputcount = doc2.all.length;

            this.msg = "\n frames：" + frames + " \n doc3_iframe_lenght：" + doc3_iframe_lenght;
            util.log_to(this.Debug_config, this.methodName, this.msg);

            
            /*******  第一层--------查找当前首层网页里》所有（IFrame）*******/
            if (doc3 != null)
            {

                foreach (IHTMLElement iframe_obj in doc3.getElementsByTagName("iframe"))
                {

                    //判断-当前IFrame是否匹配 >iframe.src( Contains == Like)
                    string iframe_obj_src = iframe_obj.getAttribute("src").ToString();
                    this.msg = "IFrame.name=" + iframe_obj.getAttribute("name").ToString();
                    util.log_to(this.Debug_config, this.methodName, this.msg);

                    if (iframe_obj_src.Contains(like_url))
                    {
                        //取出匹配到的(网页文档-Document2)
                        mshtml.IHTMLFrameBase2 FrameBase2 = (mshtml.IHTMLFrameBase2)iframe_obj;
                        mshtml.IHTMLDocument2 Document2 = FrameBase2.contentWindow.document;

                        string iframe_name = iframe_obj.getAttribute("name").ToString();
                        this.msg = this.methodName + " \nLv5>找到匹配的URL-IFrame.name=>" + iframe_name + "\nLv5>*****找到匹配的URL-IFrame.src=>" + iframe_obj_src + " 包含匹配字符〉" + like_url;
                        util.log_to(this.Debug_config, this.methodName, this.msg);

                        //@"\url_remote_js.txt"--获取配置文件中JS的URL
                        string url_remote_js = util.getFilename("url_remote_js");

                        util.add_doc_element_test(Document2, url_remote_js);

                        //string iframe_obj_name = iframe_obj.getAttribute("name").ToString();
                        //iframe_name_arrylist.Add(iframe_obj_name);//将找到的Iframe.name 属性保存起来
                        /***************************/
                    }
                    else
                    {
                        this.msg = "\nLv1>*URL-IFrame.src=> " + iframe_obj_src + " 不包含匹配字符〉" + like_url;
                        util.log_to(this.Debug_config, this.methodName, this.msg);
                    }


                }
            }
            else
            {
                
                this.msg = "doc3=null";
                util.log_to(this.Debug_config, this.methodName, this.msg);
            }



        }

        /*自定义挂载函数*****************************************************************************************************************************/
        /// <summary>
        /// 时间：2016年9月14日13:07:25
        /// 自定义挂载函数
        /// 1.加载文件内容
        ///   找到匹配的网址URL
        ///   再将些参数传给函数－OnDocumentComplete_query_iframe_like_url（like_url）处理
        /// </summary>
        /// <param name="pDisp"></param>
        /// <param name="URL"></param>
        public void OnDocumentComplete_start(object display, ref object URL)
        {
            this.Debug_config = util.getFilename("debug_config");/*初始化：读取Debug_Msg 配置是否==‘开’*/
            /**日志**/
            this.methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;//程序执行位置代码块的方法名
            this.msg = methodName + "进入此方法！";
            /***/
            util.log_to(this.Debug_config, this.methodName, this.msg);
            /**日志**/
    

            mshtml.HTMLDocument doc = (mshtml.HTMLDocument)webBrowser.Document;
            mshtml.IHTMLDocument2 doc2 = (mshtml.IHTMLDocument2)webBrowser.Document;
            mshtml.IHTMLDocument3 doc3 = (mshtml.IHTMLDocument3)webBrowser.Document;
            mshtml.IHTMLDocument4 doc4 = (mshtml.IHTMLDocument4)webBrowser.Document;
            mshtml.IHTMLDocument5 doc5 = (mshtml.IHTMLDocument5)webBrowser.Document;

            string strUrl = URL.ToString();
            //this.strFilterKeys = this.get_file_str();
            this.msg = "当前网址=" + strUrl;
            util.log_to(this.Debug_config, this.methodName, this.msg);
            string Directory_path = System.Reflection.Assembly.GetExecutingAssembly().Location;

            /*********************(读取分隔符)*********************/
            // @"\url_fgf.txt";//读取分隔符
            string url_fgf_str = util.getFilename("url_fgf");
            //System.Windows.Forms.MessageBox.Show("分隔符-url_fgf_str.ToCharArray()=" + url_fgf_str.ToCharArray());


            /*********************(读取URL)*********************/
            //@"\url.txt";//读取URL
            string urls = util.getFilename("url");
            this.msg = "读取配置文件中》URL=" + urls;
            util.log_to(this.Debug_config, this.methodName, this.msg);

            /*********************(读取URL)*********************/
            string[] url_list = urls.Split(url_fgf_str.ToCharArray());

            foreach (string url in url_list)
            {
                if (strUrl.Contains(url))//判断-是否包含-URL网址
                {

                    int frames = doc.parentWindow.frames.length;
                    int doc3_iframe_lenght = doc3.getElementsByTagName("iframe").length; //iframe 数量   //int doc2_iframe_lenght=doc2.getElementsByTagName("iframe").length; //iframe 数量 e不能正常出数
                    
                    this.msg = methodName + " \nLv5>匹配到网址-url：>" + url + "关键字的地址!\n frames=" + frames + "\n doc3_iframe_lenght=" + doc3_iframe_lenght;
                    util.log_to(this.Debug_config, this.methodName, this.msg);


                    //将当前网址URL-匹配到-文件中URL传给相关函数处理
                    this.OnDocumentComplete_query_iframe_like_url(url);

                }
                else
                {
                    this.msg = "没有找到匹配项=" + url;
                    util.log_to(this.Debug_config, this.methodName, this.msg);
                    //this.webBrowser.StatusText = strUrl;
                }
            }
            
            

        }
        /******************************************************************************************************************************/








 
        /// <summary>
        /// 注册表位置
        /// </summary>
       public static string BHOKEYNAME = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";
 
        //[ComRegisterFunction]
        //public static void RegisterBHO(Type type)
        //{
        //     RegistryKey registryKey = Registry.LocalMachine. OpenSubKey (BHOKEYNAME, true);

        //     if (registryKey == null )
        //     registryKey = Registry .LocalMachine .CreateSubKey(BHOKEYNAME);

        //     String guid = type.GUID.ToString("B");
        //     RegistryKey ourKey = registryKey.OpenSubKey (guid);

        //     if(ourKey == null) 
        //     ourKey = registryKey.CreateSubKey (guid);

        //     ourKey.SetValue("Alright", 1);
        //     registryKey.Close();
        //     ourKey.Close();
        //  }


       /// <summary>
       /// 注册FilterIEHelper。
       /// </summary>
       /// <param name="t"></param>
       [ComRegisterFunction]
       public static void RegisterBHO(Type t)
       {
           RegistryKey key = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);

           if (key == null)
               key = Registry.LocalMachine.CreateSubKey(BHOKEYNAME);

           string guidString = t.GUID.ToString("B").ToUpper();
           RegistryKey bhoKey = key.OpenSubKey(guidString);

           if (bhoKey == null)
               bhoKey = key.CreateSubKey(guidString);

           key.Close();
           bhoKey.Close();
       }

       /// <summary>
       /// 取消注册FilterIEHelper。
       /// </summary>
       /// <param name="t"></param>
       [ComUnregisterFunction]
       public static void UnRegisterBHO(Type t)
       {
           RegistryKey key = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);
           string guidString = t.GUID.ToString("B").ToUpper();

           if (key != null)
               key.DeleteSubKey(guidString, false);
       }



       /*---------------------------------------------------------------------------------------------------------------*/
       public int GetSite(ref Guid guid, out IntPtr ppvSite)
       {
           IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
           int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
           Marshal.Release(punk);

           return hr;
       }

       /*-挂载IE事件--------------------------------------------------------------------------------------------------------------*/
        /// <summary>
       /// 挂载IE事件
        /// </summary>
        /// <param name="site"></param>
        /// <returns></returns>
       public int SetSite(object site)
       {
           string debug = util.getFilename("debug");/*初始化：读取Debug_Msg 配置是否==‘开’*/
           /**日志**/
           this.methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;//程序执行位置代码块的方法名
           this.msg = methodName + "进入此方法！";
           /***/
           util.log_to(debug, this.methodName, this.msg);
           /**日志**/
           if (site != null)
           {
               webBrowser = (WebBrowser)site;
               if (debug.Equals("debug"))
               {
                   webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDucumentComplete);//测试使用挂载函数
               }else
               {
                   webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete_start);//正式环境挂载函数
               }
               //webBrowser.BeforeNavigate2 += new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
           }
           else
           {
               if (debug.Equals("debug"))
               {
                   webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDucumentComplete);//测试使用挂载函数
               }
               else
               {
                   webBrowser.DocumentComplete -= new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete_start);//正式环境挂载函数
               }
               //webBrowser.BeforeNavigate2 -= new DWebBrowserEvents2_BeforeNavigate2EventHandler(this.OnBeforeNavigate2);
               webBrowser = null;
           }
           return 0;
       }//end public int SetSite(object site)
    }
}