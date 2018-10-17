// xlltemplate.cpp
#include <cmath>
#include "xlltemplate.h"

using namespace xll;

// Information Excel needs to register add-in.
AddIn xai_function(
    Function(XLL_DOUBLE, L"?xll_function", L"XLL.FUNCTION")
    .Arg(XLL_DOUBLE, L"x", L"is the first double argument.")
    .Category(CATEGORY)
    .FunctionHelp(L"Help on XLL.FUNCTION goes here.")
);
double WINAPI xll_function(double x)
{
#pragma XLLEXPORT
    return exp(x);
}
