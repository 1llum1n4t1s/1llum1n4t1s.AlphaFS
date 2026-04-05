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

namespace Alphaleonis.Win32.Network
{
   /// <summary>DFS 名前空間内の DFS ルートまたはリンクターゲット、または DFS クライアントが管理するキャッシュからの情報を含みます。
   /// <para>このクラスは継承できません。</para>
   /// </summary>
   [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Dfs")]
   public sealed class DfsStorageInfo
   {
      #region コンストラクター

      /// <summary>DFS ルートまたはリンクターゲットのラッパーとして機能する <see cref="DfsStorageInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      public DfsStorageInfo()
      {
      }

      /// <summary>DFS ルートまたはリンクターゲットのラッパーとして機能する <see cref="DfsStorageInfo"/> クラスの新しいインスタンスを初期化します。</summary>
      /// <param name="structure">初期化された <see cref="NativeMethods.DFS_STORAGE_INFO_1"/> インスタンス。</param>
      internal DfsStorageInfo(NativeMethods.DFS_STORAGE_INFO_1 structure)
      {
         ServerName = structure.ServerName;
         ShareName = structure.ShareName;

         State = structure.State;

         TargetPriorityClass = structure.TargetPriority.TargetPriorityClass;
         TargetPriorityRank = structure.TargetPriority.TargetPriorityRank;
      }

      #endregion // コンストラクター

      #region メソッド

      /// <summary>DFS ルートターゲットまたはリンクターゲットの共有名。</summary>
      /// <returns>このインスタンスを表す文字列。</returns>
      public override string ToString()
      {
         return ShareName;
      }

      #endregion // メソッド

      #region プロパティ

      /// <summary>DFS ルートターゲットまたはリンクターゲットのサーバー名。</summary>
      public string ServerName { get; private set; }

      /// <summary>DFS ルートターゲットまたはリンクターゲットの共有名。</summary>
      public string ShareName { get; private set; }

      /// <summary>DFS ルートターゲットまたはリンクターゲットの <see cref="DfsStorageStates"/> 列挙型。</summary>
      public DfsStorageStates State { get; private set; }

      /// <summary>DFS ターゲットの優先度クラスとランクを含みます。</summary>
      public DfsTargetPriorityClass TargetPriorityClass { get; private set; }

      /// <summary>ターゲットの優先度ランク値を指定します。デフォルト値は 0 で、優先度クラス内の最高優先度ランクを示します。</summary>
      public int TargetPriorityRank { get; private set; }

      #endregion // プロパティ
   }
}
