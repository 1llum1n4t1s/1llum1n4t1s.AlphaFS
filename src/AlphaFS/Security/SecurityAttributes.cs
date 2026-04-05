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

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security.AccessControl;

namespace Alphaleonis.Win32.Security
{
   internal static partial class NativeMethods
   {
      /// <summary>SECURITY_ATTRIBUTES ネイティブWin32構造体を表すクラス。
      /// SECURITY_ATTRIBUTES構造体はオブジェクトのセキュリティ記述子を含み、この構造体を指定して取得されたハンドルが継承可能かどうかを指定します。
      /// この構造体は、CreateFile、CreatePipe、CreateProcess、RegCreateKeyEx、RegSaveKeyExなどのさまざまな関数によって作成されるオブジェクトのセキュリティ設定を提供します。
      /// </summary>
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal sealed class SecurityAttributes : IDisposable
      {
         // StructLayout属性を削除するとエラーが発生します。


         [MarshalAs(UnmanagedType.U4)]
         private int _length;

         private readonly SafeGlobalMemoryBufferHandle _securityDescriptor;


         public SecurityAttributes(ObjectSecurity securityDescriptor)
         {
            var safeBuffer = ToUnmanagedSecurityAttributes(securityDescriptor);

            _length = safeBuffer.Capacity;
            _securityDescriptor = safeBuffer;
         }


         public SecurityAttributes(ObjectSecurity securityDescriptor, bool inheritHandle) : this(securityDescriptor)
         {
            InheritHandle = inheritHandle;
         }


         public bool InheritHandle { get; set; }


         /// <summary>ObjectSecurityインスタンスをアンマネージメモリにマーシャリングします。</summary>
         /// <returns>マーシャリングされたセキュリティ記述子を含むセーフハンドル。</returns>
         /// <param name="securityDescriptor">セキュリティ記述子。</param>
         [SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic")]
         private static SafeGlobalMemoryBufferHandle ToUnmanagedSecurityAttributes(ObjectSecurity securityDescriptor)
         {
            if (null == securityDescriptor)
            {
               return new SafeGlobalMemoryBufferHandle();
            }


            var src = securityDescriptor.GetSecurityDescriptorBinaryForm();
            var safeBuffer = new SafeGlobalMemoryBufferHandle(src.Length);

            try
            {
               safeBuffer.CopyFrom(src, 0, src.Length);
               return safeBuffer;
            }
            catch
            {
               safeBuffer.Close();
               throw;
            }
         }


         public void Dispose()
         {
            if (null != _securityDescriptor && !_securityDescriptor.IsClosed)
            {
               _securityDescriptor.Close();
            }
         }
      }
   }
}
