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
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace Alphaleonis.Win32.Network
{
   /// <summary>Represents a network on the local machine. It can also represent a collection of network connections with a similar network signature.</summary>
   [Serializable]
   public class NetworkInfo : IEquatable<NetworkInfo>, IDisposable
   {
      #region Private Fields

      [NonSerialized]
      private NativeMethods.NetworkWrapper _network;

      #endregion // Private Fields


      #region Constructors

      internal NetworkInfo(NativeMethods.NetworkWrapper network)
      {
         _network = network;
      }

      #endregion // Constructors


      #region IDisposable

      /// <summary>Releases the underlying COM reference.</summary>
      public void Dispose()
      {
         _network?.Dispose();
         _network = null;
      }

      #endregion // IDisposable


      #region Private Helpers

      private void ThrowIfDisposed()
      {
         if (null == _network)
            throw new ObjectDisposedException(GetType().FullName);
      }

      #endregion // Private Helpers


      #region Properties

      /// <summary>Gets the category of a network. The categories are trusted, untrusted, or authenticated. This value of this property is not cached.</summary>
      public NetworkCategory Category
      {
         get { ThrowIfDisposed(); return _network.GetCategory(); }
      }


      /// <summary>Gets the network connections for the network. This value of this property is not cached.</summary>
      public IEnumerable<NetworkConnectionInfo> Connections
      {
         get
         {
            ThrowIfDisposed();

            // Eagerly wrap all COM wrappers into NetworkConnectionInfo objects (which have finalizers)
            // to prevent COM reference leaks if the caller partially enumerates.
            var connections = _network.GetNetworkConnections();
            var result = new List<NetworkConnectionInfo>();

            try
            {
               foreach (var connection in connections)
                  result.Add(new NetworkConnectionInfo(connection));
            }
            catch
            {
               foreach (var item in result)
                  item.Dispose();

               throw;
            }

            return result;
         }
      }


      /// <summary>Gets the local date and time when the network was connected. This value of this property is not cached.</summary>
      public DateTime ConnectionTime
      {
         get { return ConnectionTimeUtc.ToLocalTime(); }
      }


      /// <summary>Gets the date and time when the network was connected. This value of this property is not cached.</summary>
      public DateTime ConnectionTimeUtc
      {
         get
         {
            ThrowIfDisposed();

            uint unused1, unused2;

            _network.GetTimeCreatedAndConnected(out unused1, out unused2, out var low, out var high);
            
            long time = high;

            // Shift the day info into the high order bits.
            time <<= 32;
            time |= low;

            return DateTime.FromFileTimeUtc(time);
         }
      }


      /// <summary>Gets the connectivity state of the network. This value of this property is not cached.</summary>
      /// <remarks>Connectivity provides information on whether the network is connected, and the protocols in use for network traffic.</remarks>
      public ConnectivityStates Connectivity
      {
         get { ThrowIfDisposed(); return _network.GetConnectivity(); }
      }


      /// <summary>Gets the local date and time when the network was created. This value of this property is not cached.</summary>
      public DateTime CreationTime
      {
         get { return CreationTimeUtc.ToLocalTime(); }
      }


      /// <summary>Gets the date and time when the network was created. This value of this property is not cached.</summary>
      public DateTime CreationTimeUtc
      {
         get
         {
            ThrowIfDisposed();

            uint unused1, unused2;

            _network.GetTimeCreatedAndConnected(out var low, out var high, out unused1, out unused2);

            long time = high;

            // Shift the value into the high order bits.
            time <<= 32;
            time |= low;

            return DateTime.FromFileTimeUtc(time);
         }
      }


      /// <summary>Gets a description for the network. This value of this property is not cached.</summary>
      public string Description
      {
         get { ThrowIfDisposed(); return _network.GetDescription(); }

         // Should we allow this in AlphaFS?
         //private set { _network.SetDescription(value); }
      }


      /// <summary>Gets the domain type of the network. This value of this property is not cached.</summary>
      /// <remarks>The domain indictates whether the network is an Active Directory Network, and whether the machine has been authenticated by Active Directory.</remarks>
      public DomainType DomainType
      {
         get { ThrowIfDisposed(); return _network.GetDomainType(); }
      }


      /// <summary>Gets a value that indicates whether there is network connectivity. This value of this property is not cached.</summary>
      public bool IsConnected
      {
         get { ThrowIfDisposed(); return _network.IsConnected; }
      }


      /// <summary>Gets a value that indicates whether there is Internet connectivity. This value of this property is not cached.</summary>
      public bool IsConnectedToInternet
      {
         get { ThrowIfDisposed(); return _network.IsConnectedToInternet; }
      }


      /// <summary>Gets the name of the network. This value of this property is not cached.</summary>
      public string Name
      {
         get { ThrowIfDisposed(); return _network.GetName(); }

         // Should we allow this in AlphaFS?
         //private set { _network.SetName(value); }
      }


      /// <summary>Gets a unique identifier for the network. This value of this property is not cached.</summary>
      [SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "ID")]
      public Guid NetworkId
      {
         get { ThrowIfDisposed(); return _network.GetNetworkId(); }
      }

      #endregion // Properties


      #region Methods
      
      /// <summary>Returns storage device as: "VendorId ProductId DeviceType DeviceNumber:PartitionNumber".</summary>
      /// <returns>A string that represents this instance.</returns>
      public override string ToString()
      {
         var description = !Utils.IsNullOrWhiteSpace(Description) && !Equals(Name, Description) ? " (" + Description + ")" : string.Empty;

         return null != Name ? string.Format(CultureInfo.CurrentCulture, "{0}{1}, {2}", Name, description, Category) : GetType().Name;
      }


      /// <summary>Serves as a hash function for a particular type.</summary>
      /// <returns>A hash code for the current Object.</returns>
      public override int GetHashCode()
      {
         return NetworkId.GetHashCode();
      }
      

      /// <summary>Determines whether the specified Object is equal to the current Object.</summary>
      /// <param name="other">Another <see cref="NetworkInfo"/> instance to compare to.</param>
      /// <returns><c>true</c> if the specified Object is equal to the current Object; otherwise, <c>false</c>.</returns>
      public bool Equals(NetworkInfo other)
      {
         return null != other && GetType() == other.GetType() &&
                Equals(NetworkId, other.NetworkId);
      }


      /// <summary>Determines whether the specified Object is equal to the current Object.</summary>
      /// <param name="obj">Another object to compare to.</param>
      /// <returns><c>true</c> if the specified Object is equal to the current Object; otherwise, <c>false</c>.</returns>
      public override bool Equals(object obj)
      {
         var other = obj as NetworkInfo;

         return null != other && Equals(other);
      }


      /// <summary>Implements the operator ==</summary>
      /// <param name="left">A.</param>
      /// <param name="right">B.</param>
      /// <returns>The result of the operator.</returns>
      public static bool operator ==(NetworkInfo left, NetworkInfo right)
      {
         return ReferenceEquals(left, null) && ReferenceEquals(right, null) ||
                !ReferenceEquals(left, null) && !ReferenceEquals(right, null) && left.Equals(right);
      }


      /// <summary>Implements the operator !=</summary>
      /// <param name="left">A.</param>
      /// <param name="right">B.</param>
      /// <returns>The result of the operator.</returns>
      public static bool operator !=(NetworkInfo left, NetworkInfo right)
      {
         return !(left == right);
      }

      #endregion // Methods
   }
}
