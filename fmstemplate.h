// fmstemplate.h - Platform independent code
#pragma once
#define _USE_MATH_DEFINES
#include <math.h>
#include <stdexcept>
#include <utility>

#ifndef M_SQRT2PI
#define M_SQRT2PI  2.50662827463100050240
#endif

namespace fms {

    struct normal {
        static double pdf(double x, double mu = 0, double sigma = 1)
        {
            auto x_ = (x - mu)/sigma;

            return exp(-x_*x_/2)/M_SQRT2PI;
        }
        static double cdf(double x, double mu = 0, double sigma = 1)
        {
            auto x_ = (x - mu)/sigma;

            return (1 + erf(x_/M_SQRT2))/2;
        }
        static double inv(double p)
        {
            if (p <= 0 || p >= 1) {
                throw std::domain_error("normal::inv argument must be in the interval (0,1)");
            }
            // logistic(x) = 1/(1 + exp(-alpha * x))
            // logistic'(0) = cdf'(0) = pdf(0)
            static double alpha = 4 / M_SQRT2PI;

            // inverse logistic function is initial guess
            double x_, x = -log(1/p - 1) / alpha;

            do {
                x_ = x - (cdf(x) - p) / pdf(x);
                std::swap(x_, x);
            } while (fabs(x_ - x) > 2*std::numeric_limits<double>::epsilon());

            return x_;
        }
    };


}
