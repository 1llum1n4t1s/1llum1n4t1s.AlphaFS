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

namespace Alphaleonis.Win32.Security
{
   internal static partial class NativeMethods
   {
      /// <summary>TOKEN_ELEVATION_TYPE列挙型は、GetTokenInformation関数によって照会されるトークンの昇格タイプを示します。</summary>
      /// <remarks>
      /// <para>サポートされる最小クライアント: Windows Vista [デスクトップアプリのみ]</para>
      /// <para>サポートされる最小サーバー: Windows Server 2008 [デスクトップアプリのみ]</para>
      /// </remarks>
      internal enum TOKEN_ELEVATION_TYPE
      {
         ///// <summary>The token does not have a linked token: UAC is disabled or the process is started by a standard User (not a member of the Administrators group).</summary>
         //TokenElevationTypeDefault = 1,

         /// <summary>トークンは昇格されたトークンです: UACが有効でユーザーが昇格されています。</summary>
         TokenElevationTypeFull = 2,

         ///// <summary>The token is a limited token: UAC is enabled but User is not elevated.</summary>
         //TokenElevationTypeLimited = 3
      }
   }
}
