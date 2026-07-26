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

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;

namespace Alphaleonis.Win32.Filesystem
{
   /// <summary>ファイルの作成、コピー、削除、移動、およびオープンのための静的メソッドを提供し、<see cref="System.IO.FileStream"/>オブジェクトの作成を支援します。
   ///   <para>このクラスは継承できません。</para>
   /// </summary>
   [SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling")]   
   public static partial class File
   {
      internal static void ThrowIOExceptionIfFsoExist(KernelTransaction transaction, bool isFolder, string fsoPath, PathFormat pathFormat)
      {
         if (ExistsCore(transaction, isFolder, fsoPath, pathFormat))

         {
            throw new IOException(string.Format(CultureInfo.InvariantCulture, "({0}) {1}", Win32Errors.ERROR_ALREADY_EXISTS,

               // HResult には生の Win32 コードではなく HRESULT へ変換した値を入れる。0x000000B7 は
               // severity bit が 0 なので HRESULT の規約上「成功」を意味してしまい、
               // Marshal.GetExceptionForHR や ex.HResult < 0 による判定が誤動作する。
               // リポジトリ内の他の全経路 (NativeError / AlreadyExistsException) と同じ変換に揃える。
               string.Format(CultureInfo.InvariantCulture, isFolder ? Resources.Target_File_Is_A_Directory : Resources.Target_Directory_Is_A_File, fsoPath)), Win32Errors.GetHrFromWin32Error(Win32Errors.ERROR_ALREADY_EXISTS));
         }
      }
   }
}
