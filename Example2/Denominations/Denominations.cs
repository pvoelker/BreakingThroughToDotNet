// Breaking Through to .NET - Example 2 - COM interop collections
// March 2014
// Paul Voelker

// NOTES:
// --- This example demonstrates both weakly and strongly typed collections.
// --- 'Register for COM interop' is enabled under build settings.  This will
//     create a type library (TLB) file in the output folder.
// --- There is a post build event to copy the type library (TLB) file
//     to a common location so the test application can use it readily.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using COMHelpers;

namespace DenominationsBusinessLogic
{
    #region Denomination Data

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("80DC00F0-A14A-4958-BD4D-6C1739F8420A")]  // MUST be unqiue
    public interface IDenomination
    {
        string Name { get; set; }
        decimal Value { get; set; }
        bool IsCoin { get; set; }
    }

    [ComVisible(true)]
    [ComDefaultInterface(typeof(IDenomination))]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A5156097-480C-40C6-86BE-88983919C015")]  // MUST be unqiue
    public class Denomination : IDenomination
    {
        public Denomination(string name, decimal value, bool isCoin)
        {
            Name = name;
            Value = value;
            IsCoin = isCoin;
        }
        public string Name { get; set; }
        public decimal Value { get; set; }
        public bool IsCoin { get; set; }
    }

    #endregion

    #region Strongly Typed Denomination Collection

    // NOTE: If you are planning on exposing a strongly typed collection, I
    //       would recommend using a generic collection
    //       (System.Collections.Generic) as a base.  A non-generic collection
    //       has been used here for demonstration purposes only.

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("32F36D30-2171-4BBB-BB9B-9D0CD527C9C6")]  // MUST be unqiue
    public interface IDenominationListCOM
    {
        int Count { get; }

        void Clear();

        void Add(Denomination item);
        void Insert(int index, Denomination item);

        void RemoveAt(int index);
        void Remove(Denomination item);

        Denomination GetItem(int index);
        int IndexOf(Denomination item);
        bool Contains(Denomination item);
    }

    [ComVisible(true)]
    [ComDefaultInterface(typeof(IDenominationListCOM))]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("38905D32-5CE4-4859-B3C8-0DEA58EC4BD2")]  // MUST be unqiue
    public class DenominationListCOM :
        GenericListImplementationCOM<Denomination>, IDenominationListCOM { }

    #endregion

    #region Denomination Manager

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("867D70E8-43CF-4613-94C4-874450F6336B")]  // MUST be unqiue
    public interface IDenominationManager
    {
        ArrayList DenominationsWeakColl { get; } // Weakly typed collection returned
        IDenominationListCOM DenominationsStrongColl { get; } // Strongly typed collection returned

        #region Test functions

        void Test_GenerateRandomData(int size, int nameSize);
        int Test_TimeEnumerateDenominations();

        #endregion
    }

    [ComVisible(true)]
    [ComDefaultInterface(typeof(IDenominationManager))]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("3931DC97-1AA2-448C-BCB6-705233DAF538")]  // MUST be unqiue
    public class DenominationManager : IDenominationManager
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern int GetTickCount();

        DenominationListCOM denomList = new DenominationListCOM();

        public DenominationManager()
        {
            // Load collection up with data
            denomList.Add(new Denomination("Pennies", 0.01m, true));
            denomList.Add(new Denomination("Nickels", 0.05m, true));
            denomList.Add(new Denomination("Dimes", 0.10m, true));
            denomList.Add(new Denomination("Quarters", 0.25m, true));
            denomList.Add(new Denomination("Dollars", 1.00m, false));
            denomList.Add(new Denomination("Fives", 5.00m, false));
            denomList.Add(new Denomination("Tens", 10.00m, false));
            denomList.Add(new Denomination("Twenties", 20.00m, false));
        }

        public ArrayList DenominationsWeakColl { get { return denomList; } }
        public IDenominationListCOM DenominationsStrongColl { get { return denomList; } }

        #region Test functions

        public void Test_GenerateRandomData(int size, int nameSize)
        {
            denomList.Clear();

            Random rnd = new Random();
            const string SAMPLE_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            for (int i = 0; i < size; i++)
            {
                char[] buffer = new char[nameSize];
                for (int i2 = 0; i2 < nameSize; i2++)
                {
                    buffer[i2] = SAMPLE_CHARS[rnd.Next(SAMPLE_CHARS.Length)];
                }

                denomList.Add(new Denomination(new String(buffer), 0.00m, false));
            }        
        }
        [MethodImpl(MethodImplOptions.NoOptimization)]
        public int Test_TimeEnumerateDenominations()
        {
            int startTickCount = GetTickCount();

            for (int i = 0; i < denomList.Count; i++)
            {
                Denomination denom = denomList.GetItem(i);
            }

            int endTickCount = GetTickCount();

            int deltaTickCount = endTickCount - startTickCount;

            return deltaTickCount;
        }

        #endregion
    }

    #endregion
}
