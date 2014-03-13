#ifndef __COM_DECIMAL_H__
#define __COM_DECIMAL_H__

#ifndef __wtypes_h__
#error "You must include wtypes.h before you include decimal.h"
#endif // __wtypes_h__

#ifndef _OLEAUTO_H_
#error "You must include oleauto.h before you can include decimal.h"
#endif // _OLEAUTO_H_

// Change the following block if you're working with MFC, ATL, or whatever else than libc
// BEGIN
#include <assert.h>
#ifdef _DEBUG
#define verify(x)  assert(x)
#else
#define verify(x)  (x)
#endif // _DEBUG
// END

// Nevermind the const_casts scattered all over. It just happens that the OLE/COM helper 
// functions such as VarDecFromStr() is braindead. They should be accepting LPCOLESTR 
// IMHO, but they accept only OLECHAR* when they're not changing the contents of the 
// string being passed in. What rocket scientist designed this interface!?

namespace com {
	/*
	 * C++ wrapper class for the OLE Automation type DECIMAL. Provides all the VarDecXXX() functions through
	 * a C++ class and function interface. Objects of type Decimal look and behave like "ordinary ints".
	 */
	class Decimal : public DECIMAL {
	public:
		// Constructors
		Decimal(const DECIMAL& decValue) throw() {
			::CopyMemory(this, &decValue, sizeof(DECIMAL));
		}

		// Default to 0
		Decimal() throw() {
			DECIMAL_SETZERO(*this);
		}

		Decimal(char nValue) throw() {
			verify(SUCCEEDED(::VarDecFromI1(nValue, this)));
		}

		Decimal(short nValue) throw() {
			verify(SUCCEEDED(::VarDecFromI2(nValue, this)));
		}

		Decimal(long nValue) throw () {
			verify(SUCCEEDED(::VarDecFromI4(nValue, this)));
		}

		Decimal(unsigned char nValue) throw() {
			verify(SUCCEEDED(::VarDecFromUI1(nValue, this)));
		}

		Decimal(unsigned short nValue) throw() {
			verify(SUCCEEDED(::VarDecFromUI2(nValue, this)));
		}

		Decimal(unsigned long nValue) throw () {
			verify(SUCCEEDED(::VarDecFromUI4(nValue, this)));
		}

		Decimal(float fValue) throw() {
			verify(SUCCEEDED(::VarDecFromR4(fValue, this)));
		}

		Decimal(double dValue) throw() {
			verify(SUCCEEDED(::VarDecFromR8(dValue, this)));
		}

		// Mutators
		inline bool FromString(LPCOLESTR lpszNum, LCID lcid = 0) throw() {
			HRESULT hr = ::VarDecFromStr(const_cast<OLECHAR*>(lpszNum), 0 == lcid ? ::GetSystemDefaultLCID() : lcid, 0, this);
			verify(SUCCEEDED(hr));
			return SUCCEEDED(hr);
		}

		// Operators (mutating)
		inline Decimal& operator = (char nValue) throw() {
			verify(SUCCEEDED(::VarDecFromI1(nValue, this)));
			return *this;
		}

		inline Decimal& operator = (short nValue) throw() {
			verify(SUCCEEDED(::VarDecFromI2(nValue, this)));
			return *this;
		}

		inline Decimal& operator = (long nValue) throw() {
			verify(SUCCEEDED(::VarDecFromI4(nValue, this)));
			return *this;
		}

		inline Decimal& operator = (unsigned char nValue) throw() {
			verify(SUCCEEDED(::VarDecFromI1(nValue, this)));
			return *this;
		}

		inline Decimal& operator = (unsigned short nValue) throw() {
			verify(SUCCEEDED(::VarDecFromUI2(nValue, this)));
			return *this;
		}

		inline Decimal& operator = (unsigned long nValue) throw() {
			verify(SUCCEEDED(::VarDecFromUI4(nValue, this)));
			return *this;
		}

		inline Decimal& operator = (float nValue) throw() {
			verify(SUCCEEDED(::VarDecFromR4(nValue, this)));
			return *this;
		}

		inline Decimal& operator = (double nValue) throw() {
			verify(SUCCEEDED(::VarDecFromR8(nValue, this)));
			return *this;
		}

		inline Decimal& operator += (const DECIMAL& decRHS) throw() {
			verify(SUCCEEDED(::VarDecAdd(this, const_cast<DECIMAL*>(&decRHS), this)));
			return *this;
		}

		inline Decimal& operator -= (const DECIMAL& decRHS) throw() {
			verify(SUCCEEDED(::VarDecSub(this, const_cast<DECIMAL*>(&decRHS), this)));
			return *this;
		}

		inline Decimal& operator *= (const DECIMAL& decRHS) throw() {
			verify(SUCCEEDED(::VarDecMul(this, const_cast<DECIMAL*>(&decRHS), this)));
			return *this;
		}

		inline Decimal& operator /= (const DECIMAL& decRHS) throw() {
			verify(SUCCEEDED(::VarDecDiv(this, const_cast<DECIMAL*>(&decRHS), this)));
			return *this;
		}

		inline Decimal operator++(int /* prefix operator: not used */) throw() {
			Decimal decValue(*this);
			++(*this);
			return decValue;
		}

		inline Decimal operator--(int /* prefix operator: not used */) throw() {
			Decimal decValue(*this);
			--(*this);
			return decValue;
		}

		inline Decimal& operator++() throw() {
			this->operator += (Decimal(1L));
			return *this;
		}

		inline Decimal& operator--() throw() {
			this->operator -= (Decimal(1L));
			return *this;
		}

		// Mutating functions
		inline Decimal& MakeAbsolute() throw() {
			verify(SUCCEEDED(::VarDecAbs(this, this)));
			return *this;
		}

		inline Decimal& MakeNegative() throw() {
			verify(SUCCEEDED(::VarDecNeg(this, this)));
			return *this;
		}

		inline Decimal& MakeInteger() throw() {
			verify(SUCCEEDED(::VarDecInt(this, this)));
			return *this;
		}

		inline Decimal& MakeFixed() throw() {
			verify(SUCCEEDED(::VarDecFix(this, this)));
			return *this;
		}

		// nDecimals >= 0
		inline Decimal& MakeRound(int nDecimals) throw() {
			assert(nDecimals >= 0);
			verify(SUCCEEDED(::VarDecRound(this, nDecimals, this)));
			return *this;
		}

		// Non-mutating functions
		inline Decimal Absolute() const throw() {
			return Decimal(*this).MakeAbsolute();
		}

		inline Decimal Negative() const throw() {
			return Decimal(*this).MakeNegative();
		}

		inline Decimal Integer() const throw() {
			return Decimal(*this).MakeInteger();
		}

		inline Decimal Fixed() const throw() {
			return Decimal(*this).MakeFixed();
		}

		// nDecimals >= 0
		inline Decimal Round(int nDecimals) const throw() {
			return Decimal(*this).MakeRound(nDecimals);
		}

		// Inspectors (non-mutating)
		inline bool IsNegative() const throw() {
			return DECIMAL_NEG == this->sign;
		}

		inline bool IsZero() const throw () {
			return Hi32 == 0 && Lo64 == 0L;
		}

		inline BSTR ToString(LCID lcid = 0) const throw() {
			BSTR bstr = 0;
			verify(SUCCEEDED(::VarBstrFromDec(const_cast<Decimal*>(this), 0 == lcid ? ::GetSystemDefaultLCID() : lcid, 0, &bstr)));
			return bstr;
		}
	};

	// Operators
	inline Decimal operator + (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		Decimal decResult(decLHS);
		return decResult += decRHS;
	}

	inline Decimal operator - (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		Decimal decResult(decLHS);
		return decResult -= decRHS;
	}

	inline Decimal operator * (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		Decimal decResult(decLHS);
		return decResult *= decRHS;
	}

	inline Decimal operator / (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		Decimal decResult(decLHS);
		return decResult /= decRHS;
	}

	inline bool operator < (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		HRESULT hr = ::VarDecCmp(const_cast<DECIMAL*>(&decLHS), const_cast<DECIMAL*>(&decRHS));
		assert(SUCCEEDED(hr));
		return VARCMP_LT == hr;
	}

	inline bool operator > (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		HRESULT hr = ::VarDecCmp(const_cast<DECIMAL*>(&decLHS), const_cast<DECIMAL*>(&decRHS));
		assert(SUCCEEDED(hr));
		return VARCMP_GT == hr;
	}

	inline bool operator <= (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		return !(decLHS > decRHS);
	}

	inline bool operator >= (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		return !(decLHS < decRHS);
	}

	inline bool operator == (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		HRESULT hr = ::VarDecCmp(const_cast<DECIMAL*>(&decLHS), const_cast<DECIMAL*>(&decRHS));
		assert(SUCCEEDED(hr));
		return VARCMP_EQ == hr;
	}

	inline bool operator != (const DECIMAL& decLHS, const DECIMAL& decRHS) throw() {
		return !(decLHS == decRHS);
	}


	// Very specialized function for rounding hundreds. Consider a currency X, which smallest coin is
	// 50 cents. A value of 1.35 should then be rounded to 1.50. A value of 1.20 should be rounded to 
	// 1.00. That's what this function does; it rounds to nearest coins. If you feel that this is 
	// bloat, then by all means, remove it. It's just here because I'm using decimals in a POS application,
	// and I was too lazy to make a free function out of it (I really should because this is not a general
	// numeric operation, it's highly application specific!). Besides, someone might have use for it. :)

#define CENTS_PER_UNIT		100L
#define UNIT(x)				((x).Integer())
#define CENT(x)				(((x) - (UNIT(x))) * Decimal(CENTS_PER_UNIT))
#define MAKE(u, c)			((u) + c / Decimal(CENTS_PER_UNIT))
#define BETWEEN(what, x, y) ((x) <= (what) && (what) < (y))

	// Param: decThis          represents the number you want to round
	// Param: nSmallestCoin    an integer in cents (cent = hundreth of a currency), specifying the smallest coin. 
	Decimal RoundToSmallestCoin(const Decimal& decThis, unsigned long nSmallestCoin) {
		// Hmm.. smallest coin is infinitely small? Well, we're done then I guess...
		if(!nSmallestCoin) 
			return decThis;

		// This algorithm is designed for absolute values. Negate, calculate, and negate again..
		if(decThis.IsNegative())
			return RoundToSmallestCoin(decThis.Negative(), nSmallestCoin).MakeNegative();

		Decimal minimum, maximum;
		Decimal unit = UNIT(decThis); 
		const Decimal cent =  CENT(decThis);
		const Decimal centsPerUnit(CENTS_PER_UNIT);
		const Decimal decTwo(2L);
		const Decimal decSmallestCoin(nSmallestCoin);

		minimum = 0L;
		maximum = decSmallestCoin;

		do 
		{
			Decimal middle = (maximum - minimum) / Decimal(2L);

			if(BETWEEN(cent, minimum, minimum + middle))
				return MAKE(unit, minimum);

			else if(BETWEEN(cent, minimum + middle, maximum))
				return MAKE(unit, maximum);


			minimum = maximum;
			maximum += decSmallestCoin;
		} while(minimum < centsPerUnit);

		++unit; // Overflow
		return unit;
	}
#undef CENTS_PER_UNIT
#undef UNIT
#undef CENT
#undef MAKE
#undef BETWEEN
}

#undef verify

#endif // __COM_DECIMAL_H__