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

AddIn xai_what(
    Function(XLL_LPOPER, L"?xll_what", L"XLL.WHAT")
    .Arg(XLL_LPOPER, L"oper", L"is an reference to a cell or array of cells")
    .Category(CATEGORY)
    .FunctionHelp(L"What is the argument?")
);
LPOPER WINAPI xll_what(LPOPER po)
{
#pragma XLLEXPORT
    static OPER o;

    switch (po->xltype) {
    case xltypeNum:
        o = L"number: ";
        o = o & *po;
        break;
    case xltypeStr:
        o = L"string: ";
        o = o & *po;
        break;
    case xltypeBool:
        o = L"boolean: ";
        o = o & *po;
        break;
    case xltypeRef:
        o = L"reference: ";
        o = o & Excel(xlfReftext, *po, OPER(true));
        break;
    case xltypeErr:
        o = L"error: ";
        o = o & *po;
        break;
    case xltypeMulti:
        o = L"multi: ";
        o = o & Excel(xlfEvaluate, *po);
        break;
    case xltypeMissing:
        o = L"missing";
        o = o & *po;
        break;
    case xltypeNil: // reference to empty cell
        o = L"nil";
        o = o & *po;
        break;
    case xltypeSRef:
        o = L"simple reference: ";
        o = o & Excel(xlfReftext, *po, OPER(true));
        break;
    case xltypeInt: // this never happens
        o = L"integer: ";
        o = o & *po;
        break;
    default:
        o = L"unknown type";
    }

    return &o;
}