// Breaking Through to .NET - Example 1 - Basic COM interop
// March 2014
// Paul Voelker

// NOTES:
// --- 'Register for COM interop' is enabled under build settings.  This will
//     create a type library (TLB) file in the output folder.
// --- There is a post build event to copy the type library (TLB) file
//     to a common location so the test application can use it readily.

using System;
using System.Diagnostics;
using System.Runtime.InteropServices; // Required for COM interop
using System.Text;

namespace CurrencyConverter
{
    [ComVisible(true)]
    public enum CurrencyType
    {
        USDollar,
        Euro,
        BritishPound,
        MexicanPeso,
        JapaneseYen,
    }

    // Interface definition
    [ComVisible(true)]
    [Guid("7AEC7B1D-2475-4B34-B48D-E21314990868")] // MUST be unique (use 'GUID Generator' to create)
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] // Setting required for exposing IUnknown (early bound) interface
    public interface ICurrencyConverter
    {
        bool IsCurrencySupported(CurrencyType type);
        string GetCurrencyName(CurrencyType type);
        DateTime ExchangeRateDateTime { get; }
        int CurrenciesSupported { get; }
        decimal GetCurrentExchangeRateAgainstUSD(CurrencyType type);
        decimal ConvertCurrency(CurrencyType currencyHave,
            decimal currencyHaveAmount, CurrencyType currencyWant);
    }

    // Implementation for interface
    [ComVisible(true)]
    [ComDefaultInterface(typeof(ICurrencyConverter))] // Set to interface definition that this class inherits from
    [Guid("7E0637A7-82D3-4990-BFE3-314613B60124")] // MUST be unqiue (use 'GUID Generator' to create)
    [ClassInterface(ClassInterfaceType.None)] // Setting required for implementation class
    public class CurrencyConverter : ICurrencyConverter // Implementation class MUST inherit interface
    {
        public bool IsCurrencySupported(CurrencyType type)
        {
            bool supported = false;
            switch (type)
            {
                case CurrencyType.USDollar:
                    supported = true;
                    break;
                case CurrencyType.Euro:
                    supported = true;
                    break;
                case CurrencyType.BritishPound:
                    supported = true;
                    break;
                case CurrencyType.MexicanPeso:
                    supported = true;
                    break;
                case CurrencyType.JapaneseYen:
                    supported = true;
                    break;
            }
            return supported;
        }

        public string GetCurrencyName(CurrencyType type)
        {
            string name = "Unknown";
            switch (type)
            {
                case CurrencyType.USDollar:
                    name = "United States Dollar";
                    break;
                case CurrencyType.Euro:
                    name = "Euro";
                    break;
                case CurrencyType.BritishPound:
                    name = "British Pound Sterling";
                    break;
                case CurrencyType.MexicanPeso:
                    name = "Mexican Peso";
                    break;
                case CurrencyType.JapaneseYen:
                    name = "Japanese Yen";
                    break;
            }
            return name;
        }

        public DateTime ExchangeRateDateTime
        {
            get { return new DateTime(2014, 2, 27); }
        }

        public int CurrenciesSupported
        {
            get { return 5; }
        }

        public decimal GetCurrentExchangeRateAgainstUSD(CurrencyType type)
        {
            decimal exchangeRate = 0m;
            switch (type)
            {
                case CurrencyType.USDollar:
                    exchangeRate = 1m;
                    break;
                case CurrencyType.Euro:
                    exchangeRate = 0.73m;
                    break;
                case CurrencyType.BritishPound:
                    exchangeRate = 0.60m;
                    break;
                case CurrencyType.MexicanPeso:
                    exchangeRate = 13.30m;
                    break;
                case CurrencyType.JapaneseYen:
                    exchangeRate = 102.13m;
                    break;
            }
            return exchangeRate;
        }

        public decimal ConvertCurrency(CurrencyType currencyHave,
            decimal currencyHaveAmount, CurrencyType currencyWant)
        {
            if (IsCurrencySupported(currencyHave) == false)
            {
                ArgumentException ex =
                    new ArgumentException("Unsupported currency type", "currencyHave");
                Debug.WriteLine("***ERROR*** BreakingThroughToDotNetExample1.CurrencyConverter.ConvertCurrency Error: " + ex.ToString());
                throw ex;
            }
            else if(IsCurrencySupported(currencyWant) == false)
            {
                ArgumentException ex =
                    new ArgumentException("Unsupported currency type", "currencyWant");
                Debug.WriteLine("***ERROR*** BreakingThroughToDotNetExample1.CurrencyConverter.ConvertCurrency Error: " + ex.ToString());
                throw ex;
            }
            else
            {
                decimal currencyWantAmount =
                    currencyHaveAmount / GetCurrentExchangeRateAgainstUSD(currencyHave);
                currencyWantAmount *= GetCurrentExchangeRateAgainstUSD(currencyWant);
                return currencyWantAmount;
            }
        }
    }
}
