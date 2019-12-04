using System;
using System.Configuration;
using System.IO;

/* 添加对ocrOneNote类库的引用
 * ocrOneNote.dll
 * 
 * 在配置文件中设置ocr图像文字识别延迟时间
 * 根据图片文件大小可适当调节识别延时，确保识别完成
 <?xml version="1.0" encoding="utf-8" ?>
<configuration>
    <startup> 
        <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.5.2" />
    </startup>
	<appSettings>
	<add key="WaitTime" value="1000" />
	</appSettings>	
</configuration>
*/

namespace ocrTest
{
    class Program
    {
        static void Main(string[] args)
        {
            // 命令行检查和提示
            if (args.Length != 1 && args.Length != 2)
            {
                Console.WriteLine("\r\nOCR图像文字识别测试程序V1.0，上海因致信息科技有限公司，2018/06。");
                Console.WriteLine("\r\n命令行格式：");
                Console.WriteLine("ocrTest <图片文件名> [输出文件名]");
                Console.WriteLine("\r\n图片文件名应包含扩展名，支持jpg、gif、bmp、tif、png、emf等格式图片。");
                Console.WriteLine("输出文件为文本文件，可省略，默认输出到屏幕。");
                Console.WriteLine("\r\n用法示例：");
                Console.WriteLine("输出到屏幕：ocrTest 1.png");
                Console.WriteLine("输出到文件：ocrTest 1.png 1.txt");
                return;
            }

            // 获取ocr图像文字识别延迟时间
            int waitTime = Convert.ToInt32(ConfigurationManager.AppSettings["WaitTime"]);

            // 定义一个ocr对象
            ocrOneNote.ocr myocr = new ocrOneNote.ocr(waitTime);

            // 获取图片文件信息
            FileInfo fi = new FileInfo(args[0]);

            // ocr识别
            string strRet = myocr.OcrImg(fi);

            Console.WriteLine("\r\nOCR图像文字识别测试程序V1.0，上海因致信息科技有限公司，2018/06。");

            // 输出到屏幕
            if (args.Length == 1)
            {
                Console.WriteLine("\r\n图像文字识别结果如下：");
                Console.WriteLine("---------------------\r\n");
                Console.WriteLine(strRet);
            }

            // 输出到文本
            if (args.Length == 2)
            {
                string filePath = args[1].ToString();
                StreamWriter sw;
                if (!File.Exists(filePath))
                {
                    sw = File.CreateText(filePath);
                }
                else
                {
                    sw = File.AppendText(filePath);
                }
                sw.Write(strRet);
                sw.Close();
            }
        }
    }
}
