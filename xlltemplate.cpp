// xlltemplate.cpp
#include <cmath>
#include "fmstemplate.h"
#include "xlltemplate.h"

using namespace fms;
using namespace xll;

AddIn xai_template(
	Documentation(LR"(
This object will generate a Sandcastle Helpfile Builder project file
called <code>
)"));

// Information Excel needs to register add-in.
AddIn xai_function(
	// Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
	// Don't forget prepend a question mark to the C++ name.
	//                     |
	//                     v
    Function(XLL_LPOPER, L"?xll_function", L"XLL.FUNCTION")
	// First argument is a double called x with an argument description.
    .Arg(XLL_DOUBLE, L"x", L"is the first double argument.")
	// Paste function category.
    .Category(CATEGORY)
	// Insert Function description.
    .FunctionHelp(L"Help on XLL.FUNCTION goes here.")
	// Create entry for this function in Sandcastle Help File Builder project file.
	.Documentation(LR"(
Documentation on XLL.FUNCTION goes here.
    )")
);
// Calling convention *must* be WINAPI (aka __stdcall) for Excel.
LPOPER WINAPI xll_function(double x)
{
// Be sure to export your function.
#pragma XLLEXPORT
	static OPER result;

	try {
		ensure(x >= 0);
		result = sqrt(x); // OPER's act like Excel cells.
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = OPER(xlerr::Num);
	}

	return &result;
}

AddIn xai_normal_pdf(
    Function(XLL_DOUBLE, L"?xll_normal_pdf", L"NORMAL.PDF")
    .Arg(XLL_DOUBLE, L"x", L"is the value at which to compute the normal probability density function.")
    .Category(CATEGORY)
    .FunctionHelp(L"Compute the normal probability density function.")
    .Documentation(L"Compute the normal probability density function.")
);
double WINAPI xll_normal_pdf(double x, double mu, double sigma)
{
#pragma XLLEXPORT
    if (sigma == 0)
        sigma = 1;

    return normal::pdf(x, mu, sigma);
}

AddIn xai_normal_cdf(
    Function(XLL_DOUBLE, L"?xll_normal_cdf", L"NORMAL.CDF")
    .Arg(XLL_DOUBLE, L"x", L"is the value at which to compute the normal cumulative density function.")
    .Category(CATEGORY)
    .FunctionHelp(L"Compute the normal cumulative density function.")
    .Documentation(L"Compute the normal cumulative density function.")
);
double WINAPI xll_normal_cdf(double x, double mu, double sigma)
{
#pragma XLLEXPORT
    if (sigma == 0)
        sigma = 1;

    return normal::cdf(x, mu, sigma);
}

AddIn xai_normal_inv(
    Function(XLL_DOUBLE, L"?xll_normal_inv", L"NORMAL.INV")
    .Arg(XLL_DOUBLE, L"x", L"is the value at which to compute the inverse of the standard normal cumulative density function.")
    .Category(CATEGORY)
    .FunctionHelp(L"Compute the inverse of the standard normal cumulative density function.")
    .Documentation(L"Compute the inverse of the standard normal cumulative density function.")
);
double WINAPI xll_normal_inv(double p)
{
#pragma XLLEXPORT
    double x = std::numeric_limits<double>::quiet_NaN();

    try {
        x = normal::inv(p);
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());
    }

    return x;
}