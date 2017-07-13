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
	string name;
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
		cout << "Имя:" << name << endl;
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
	void setAll(string n, double w, double d, double h)
	{
		name = n;
		width = w;
		depth = d;
		height = h;
	}

	Paralelepiped(string n, double w, double d, double h)
	{
		setAll(n,w,d,h);
	}
	Paralelepiped()
	{
		setAll("",0,0,0);
	}
	Paralelepiped operator++()
	{
		width = width + 5;
		return *this;
	}
	Paralelepiped operator++(int)
	{
		depth++;
		return *this;
	}
	Paralelepiped operator+(Paralelepiped obj)
	{
		Paralelepiped tmp;
		tmp.name = "ОбъектС";
		tmp.width = width + obj.width;
		tmp.depth = depth + obj.depth;
		tmp.height = (height > obj.height) ? height : obj.height;
		return tmp;
	}

};
double operator-(Paralelepiped objX, Paralelepiped objY)
{
	return objX.getAmount() - objY.getAmount();
}
Paralelepiped operator--(Paralelepiped &obj)
{
	if (obj.depth > 10)
	{
		obj.depth -= 10;
	}
	else
	{
		obj.width = 30;
	}
	return obj;
}
Paralelepiped operator--(Paralelepiped &obj, int)
{
	if (obj.height > 0)
	{
		obj.height--;
	}
	else
	{
		obj.height = 1;
	}
	return obj;
}

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
		Paralelepiped::setAll("", w, d, h);
		massa = m;
	}
	MassParalelepiped(double w, double d, double h, double m) : Paralelepiped("",w, d, h)
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
	Paralelepiped objA("ОбъектА",10, 10, 10);
	objA.showAll();
	cout << endl;
	objA++;
	objA.showAll();
	objA++;
	objA.showAll();
	cout << endl;
	objA--;
	objA.showAll();
	cout << endl;
	--objA;
	objA.showAll();
	cout << endl;
	++objA;
	objA.showAll();
	cout << endl;
	Paralelepiped objB("ОбъектБ",10, 10, 10);
	objB.showAll();
	cout << endl;
	Paralelepiped objC;
	objC = objA + objB;
	objC.showAll();
	cout << "Разница в объеме: " << objA - objB << endl;
	system("pause>nul");
	return 0;
}

