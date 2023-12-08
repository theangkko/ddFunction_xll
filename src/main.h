// eefunction_xll.h - common includes
#pragma once
// Uncomment to build for pre 2007 Excel
//#define XLOPERX XLOPER
#include "xll/xll/xll.h"


double stmPTH(double p, double t);


std::string ddCheckVersion();

double ddSteamTP(double p);

double ddSteamTPH(double p, double h);

double ddSteamTPS(double p, double s);

double ddSteamTHS(double h, double s);

double ddSteamPT(double t);

double ddSteamHPT(double p, double t);

double ddSteamHPS(double p, double s);

double ddSteamHPQ(double p, double drynessX);

double ddSteamVPT(double p, double t);

double ddSteamVPH(double p, double h);

double ddSteamVPS(double p, double s);

double ddSteamRhoPT(double p, double t);

double ddSteamRhoPH(double p, double h);

double ddSteamRhoPS(double p, double s);

double ddSteamSPT(double p, double t);

double ddSteamSPH(double p, double h);

double ddSteamCpPT(double p, double t);

double ddSteamCpPH(double p, double h);

double ddSteamCvPT(double p, double t);

double ddSteamCvPH(double p, double h);

double ddSteamCvPS(double p, double s);

double ddSteamVisPT(double p, double t);

double ddSteamVisPH(double p, double h);

double ddSteamVisPS(double p, double s);

double ddSteamQPH(double p, double h);

double ddSteamQPS(double p, double s);