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
namespace Alphaleonis.Win32.Filesystem
{
   internal static partial class NativeMethods
   {
      /// <summary>1601年1月1日からの 100 ナノ秒間隔の数を表します。この構造体は 64 ビット値です。</summary>
      [Serializable]
      [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
      internal struct FILETIME
      {         
         #region Fields

         private readonly uint dwLowDateTime;
         private readonly uint dwHighDateTime;

         #endregion // Fields

         #region Methods

         /// <summary>値を long に変換します。</summary>
         public static implicit operator long(FILETIME ft)
         {
            return ft.ToLong();
         }

         /// <summary>値を long に変換します。</summary>
         [SuppressMessage("Microsoft.Naming", "CA1720:IdentifiersShouldNotContainTypeNames", MessageId = "long")]
         public long ToLong()
         {
            return NativeMethods.ToLong(dwHighDateTime, dwLowDateTime);
         }

         #endregion

         #region Equality

         #region Equals

         /// <summary>指定した Object が現在の Object と等しいかどうかを判定します。</summary>
         /// <param name="obj">比較対象の別のオブジェクト。</param>
         /// <returns>指定した Object が現在の Object と等しい場合は <c>true</c>、それ以外の場合は <c>false</c>。</returns>
         public override bool Equals(object obj)
         {
            if (null == obj || GetType() != obj.GetType())
            {
               return false;
            }

            var other = obj as FILETIME? ?? new FILETIME();

            return other.dwHighDateTime.Equals(dwHighDateTime) && other.dwLowDateTime.Equals(dwLowDateTime);
         }

         #endregion // Equals

         #region GetHashCode

         /// <summary>特定の型のハッシュ関数として機能します。</summary>
         /// <returns>現在の Object のハッシュコード。</returns>
         public override int GetHashCode()
         {
            unchecked
            {
               var hash = 17;
               hash = hash * 23 + dwHighDateTime.GetHashCode();
               hash = hash * 11 + dwLowDateTime.GetHashCode();
               return hash;
            }
         }

         #endregion // GetHashCode

         #region ==

         /// <summary>== 演算子を実装します。</summary>
         /// <param name="left">左辺の値。</param>
         /// <param name="right">右辺の値。</param>
         /// <returns>演算子の結果。</returns>
         public static bool operator ==(FILETIME left, FILETIME right)
         {
            return left.Equals(right);
         }

         #endregion // ==

         #region !=
         /// <summary>!= 演算子を実装します。</summary>
         /// <param name="left">左辺の値。</param>
         /// <param name="right">右辺の値。</param>
         /// <returns>演算子の結果。</returns>
         public static bool operator !=(FILETIME left, FILETIME right)
         {
            return !(left == right);
         }

         #endregion // !=

         #endregion // Equality
      }
   }
}
