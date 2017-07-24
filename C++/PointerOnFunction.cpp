// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <cmath>
using namespace std;
double f(double x, int n)
{
	double s = 1;
	for (int k = 1;k <= n;k++)
	{
		s *= (1 + x);
	}
	return s;
}
double g(double x, int n)
{
	double s = 1;
	for (int k = 1;k <= n;k++)
	{
		s *= x / k;
	}
	return s;
}
char h(int n)
{
	return 'A' + n;
}
char u(int n)
{
	return 'Z' - n;
}
int main()
{
	setlocale(LC_ALL, "Russian");
	double x = 2;
	int n = 3;
	double(*p)(double, int);
	char(*q)(int);
	double(*r)(double);
	p = f;
	cout << "| " << p(x, n) << " | ";
	p = g;
	cout << p(x, n) << " | ";
	q = h;
	cout << q(n) << " | ";
	q = u;
	cout << q(n) << " | ";
	r = exp;
	cout << r(x / 2) << " | ";
	r = log;
	cout << r(x) << " |\n ";
	system("pause>nul");
	return 0;
}

