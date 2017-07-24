// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <string>
using namespace std;
class Paralelepiped
{
public:
	double width;
	double depth;
	double height;
	double getAmount()
	{
		double w = width;
		double d = depth;
		double h = height;
		double a;
		a = w*d*h;
		return a;
	}
	void showAll()
	{
		cout << "Ширина: " << width << endl;
		cout << "Глубина: " << depth << endl;
		cout << "Высота: " << height << endl;
		cout << "Объем " << getAmount() << endl;
		for (int k = 1;k <= 35;k++)
		{
			cout << "-";
		}
		cout << endl;

	}
	void setAll(double w, double d, double h)
	{
		width = w;
		depth = d;
		height = h;
	}

	Paralelepiped(double w, double d, double h)
	{
		setAll(w,d,h);
	}
	Paralelepiped()
	{
		setAll(0,0,0);
	}

};
class MassParalelepiped : public Paralelepiped
{
public:
	double massa;
	double getPlot()
	{
		double w = width;
		double d = depth;
		double h = height;
		double a;
		a = (w*d*h)/massa;
		return a;
	}
	void showAll()
	{
		cout << "Ширина: " << width << endl;
		cout << "Глубина: " << depth << endl;
		cout << "Высота: " << height << endl;
		cout << "Объем: " << getAmount() << endl;
		cout << "Плотность: " << getPlot() << endl;
		for (int k = 1;k <= 35;k++)
		{
			cout << "-";
		}
		cout << endl;
	}
	void setAll(double w, double d, double h, double m)
	{
		Paralelepiped::setAll(w, d, h);
		massa = m;
	}
	MassParalelepiped(double w, double d, double h, double m) : Paralelepiped(w, d, h)
	{
		massa = m;
	}
	MassParalelepiped(): Paralelepiped()
	{
		massa = 1;
	}
};


int main()
{
	setlocale(LC_ALL, "Russian");
	Paralelepiped objA(10, 10, 10);
	Paralelepiped objB(3, 3, 3);
	Paralelepiped objC(0.1, 0.4, 2.5);
	MassParalelepiped objD;
	objD.setAll(10, 3, 2, 5);
	objA.showAll();
	cout << endl;
	objB.showAll();
	cout << endl;
	objC.showAll();
	cout << endl;
	objD.showAll();
	
	system("pause>nul");
	return 0;
}

