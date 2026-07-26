/*  Copyright (C) 2008-2018 Peter Palotas, Jeffrey Jangli, Alexandr Normuradov
 *  
 *  Permission is hereby granted, free of charge, to any person obtaining a copy 
 *  of this software and associated documentation files (the "Software"), to deal 
 *  in the Software without restriction, including without limitation the rights 
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
 *  copies of the Software, and to permit persons to whom the Software is 
 *  furnished to do so, subject to the following conditions:
 *  
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *  
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
 *  THE SOFTWARE. 
 */

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace AlphaFS.UnitTest
{
   public partial class AlphaFS_Shell32Test
   {
      // Pattern: <class>_<function>_<scenario>_<expected result>


      [TestMethod]
      public void AlphaFS_Shell32Info_GetVerbCommand_LocalAndNetwork_Success()
      {
         AlphaFS_Shell32Info_GetVerbCommand(false);
         AlphaFS_Shell32Info_GetVerbCommand(true);
      }
      
      private void AlphaFS_Shell32Info_GetVerbCommand(bool isNetwork)
      {
         // かつては ".txt は C:\Windows\System32\notepad.exe に関連付けられている" と決め打ちしていたが、
         // 現行の Windows ではストア版メモ帳や任意のエディターが既定になり得るため成立しない。
         // 特定のアプリを期待せず「実行可能ファイルを指すコマンド文字列が返ること」を検証する。


         using var tempRoot = new TemporaryDirectory(isNetwork);
         var file = tempRoot.CreateFile();

         Console.WriteLine("Input File Path: [{0}]\n", file.FullName);


         var shell32Info = Alphaleonis.Win32.Filesystem.Shell32.GetShell32Info(file.FullName);
            

         var cmd = "open";
         var result = shell32Info.GetVerbCommand(cmd);
         Console.WriteLine("\tMethod: Shell32Info.GetVerbCommand(\"{0}\")  == [{1}]", cmd, result);


         // "open" 動詞は .txt に必ず関連付けられているので、実行可能ファイルを指すコマンドが返るはず。
         Assert.IsFalse(Alphaleonis.Utils.IsNullOrWhiteSpace(result), "The \"open\" verb command is empty, but is expected not to.");
         Assert.Contains(".exe", result, StringComparison.OrdinalIgnoreCase);


         cmd = "print";
         result = shell32Info.GetVerbCommand(cmd);
         Console.WriteLine("\tMethod: Shell32Info.GetVerbCommand(\"{0}\") == [{1}]\n", cmd, result);


         // "print" 動詞は関連付けられたアプリによっては存在しない。存在する場合だけ形式を検証する。
         if (!Alphaleonis.Utils.IsNullOrWhiteSpace(result))
         {
            Assert.Contains(".exe", result, StringComparison.OrdinalIgnoreCase);
         }
      }
   }
}
