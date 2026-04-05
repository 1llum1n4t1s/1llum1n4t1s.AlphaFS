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

namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>FINDEX_INFO_LEVELS 列挙型 - FindFirstFileEx 関数と共に使用して、返されるデータの情報レベルを指定する値を定義します。</summary>
      /// <remarks>
      ///   <para>サポートされる最小クライアント: Windows XP [デスクトップアプリ | Windows ストアアプリ]</para>
      ///   <para>サポートされる最小サーバー: Windows Server 2003 [デスクトップアプリ | Windows ストアアプリ]</para>
      /// </remarks>
      internal enum FINDEX_INFO_LEVELS
      {
         /// <summary><see cref="WIN32_FIND_DATA"/> 構造体で標準的な属性セットが返されます。</summary>
         Standard = 0,

         /// <summary>FindFirstFileEx 関数は短いファイル名を問い合わせないため、全体的な列挙速度が向上します。</summary>
         /// <remarks>この値は Windows Server 2008 R2 および Windows 7 までサポートされていません。</remarks>
         Basic = 1

         ///// <summary>This value is used for validation. Supported values are less than this value.</summary>
         //MaxLevel
      }
   }
}
