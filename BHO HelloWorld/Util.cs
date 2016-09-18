using mshtml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace IE
{
    class Util
    {
        //调试信息开关
        protected string Debug_config = "off";
        protected string msg = "";
        protected string methodName = "";

        public Util() 
        {
            this.Debug_config = this.getFilename("debug_config");/*初始化：读取Debug_Msg 配置是否==‘开’*/
            string methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;//程序执行位置代码块的方法名
        }

        /*************************************************************************************************************/
        /// <summary>
        /// 时间：2016年9月14日12:41:56
        /// 传入找到的DOC文档
        /// </summary>
        /// <param name="IHTMLDocument2 Document2,string url_remote_js(文件名)" ></param>
        public void add_doc_element_test(IHTMLDocument2 Document2, string url)        
        {
            this.methodName = System.Reflection.MethodBase.GetCurrentMethod().Name;//程序执行位置代码块的方法名
            this.msg = methodName + "进入此方法！";
            this.log_to(Debug_config, methodName, this.msg);
    
            try
            {
                IHTMLElement head = (IHTMLElement)((IHTMLElementCollection)Document2.all.tags("head")).item(null, 0);
                var body = (HTMLBody)Document2.body;

                
  
                /**************************************************添加Javascript脚本******************************************/

                this.msg = methodName + "成功》加载JS：\n" + url;
                this.log_to(Debug_config, methodName, this.msg);

                IHTMLElement scriptElement = Document2.createElement("script");
                scriptElement.setAttribute("type", "text/javascript");
                scriptElement.setAttribute("src", url);
                body.insertAdjacentElement("afterBegin", scriptElement);

                // ((HTMLHeadElement)head).appendChild((IHTMLDOMNode)scriptElement);
                /**************************************************添加Javascript脚本******************************************/

                //string btnString = @"<input type='button' value='Microsoft@.net C#' onclick='FindPassword()' />";
                //body.insertAdjacentHTML("afterBegin", btnString);

            }
            catch (Exception e) 
            { 
                //System.Windows.Forms.MessageBox.Show(e.ToString());
                this.msg = methodName + e.ToString();
                this.log_to(Debug_config,methodName, this.msg); 
            }
        }
        /*************************************************************************************************************/



        /*************************************************************************************************************/
        /// <summary>
        /// 获取文件内一行字符串  ：2016年9月13日11:12:22
        /// </summary>
        /// <param name="filename_directory"></param>
        /// <returns>arry</returns>
        public string getFilename(string filename)
        {

            //string filestr = "";
            //filestr += "\n" + Environment.CurrentDirectory + "\n" + AppDomain.CurrentDomain.BaseDirectory + "\n" + AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\n" + System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;

            //filestr += "\n" + System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            //filestr += "\n" + System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            //string versionstr = System.Reflection.Assembly.GetExecutingAssembly().Location;

            //在系统服务中最好用这个方式去取路径

            string Directory_path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            Directory_path = Directory_path.Substring(0, Directory_path.LastIndexOf('\\')) + @"\" + filename + ".txt";//文件名

            FileStream fs = new FileStream(Directory_path, FileMode.Open);
            StreamReader m_streamReader = new StreamReader(fs);

            m_streamReader.BaseStream.Seek(0, SeekOrigin.Begin);
            string arry = "";
            arry = m_streamReader.ReadToEnd();

            m_streamReader.Close();
            m_streamReader.Dispose();
            fs.Close();
            fs.Dispose();

            return arry;
        }
        /*************************************************************************************************************/




        /*************************************************************************************************************/
        /// <summary>
        /// 配置文件参数1：<开log>: 开户日志输出-@C:\IE_BHO_DEBUG_LOG.txt
        /// 配置文件参数2：<开msg>：开启弹出提示窗口
        /// </summary>
        /// <param name="Debug_Msg_On_off 开log|开msg"></param>
        /// <param name="methodName"></param>
        /// <param name="debugmsg"></param>
        public void log_to(string Debug_Msg_On_off, string methodName, string debugmsg)
        {
            try
            {
                //打印到Log日志文件
                if (Debug_Msg_On_off.Equals("开log")) 
                {
                    StreamWriter sw = new StreamWriter(@"d:\IE_BHO_DEBUG_LOG.txt", true);
                    sw.WriteLine(methodName + "    >> " + debugmsg);
                    sw.Close();
                }
                
                //弹出窗口提示
                if (Debug_Msg_On_off.Equals("开msg")) 
                {
                    System.Windows.Forms.MessageBox.Show(methodName + "    >> " + debugmsg); 
                }
               


            }
            catch
            {
                // Release();
                // Marshal.ThrowExceptionForHR(E_FAIL);
            }
        }
        /*************************************************************************************************************/

    }
}
