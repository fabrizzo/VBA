// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <string>
using namespace std;
class myMoney
{
public:
	string name;
	double money;
	double rate;
	int time;
	myMoney()
	{
		name = "";
		money = 0;
		rate = 0;
		time = 0;
	}
	myMoney(string n, double m, double r, int t)
	{
		setAll(n, m, r, t);
	}
	double getMoney()
	{
		double s = money;
		for (int k = 1;k <= time;k++)
		{
			s *= (1 + rate / 100);
		}
		return s;
	}
	void showAll()
	{
		cout << "Имя: " << name << endl;
		cout << "Вклад: " << money << endl;
		cout << "Ставка (%): " << rate << endl;
		cout << "Период (лет): " << time << endl;
		cout << "Итоговая сумма: " << getMoney() << endl;
		for (int k = 1;k <= 35;k++)
		{
			cout << "-";
		}
		cout << endl;

	}
	void setAll(string n, double m, double r, int t)
	{
		name = n;
		money = m;
		rate = r;
		time = t;
	}
	myMoney operator++()
	{
		money = money+1000;
		return *this;
	}
	myMoney operator++(int)
	{
		time++;
		return *this;
	}
	myMoney operator+(myMoney obj)
	{
		myMoney tmp;
		tmp.name = "Почтальон печкин";
		tmp.money = money + obj.money;
		tmp.rate = (rate > obj.rate) ? rate : obj.rate;
		tmp.time = (time + obj.time) / 2;
		return tmp;
	}
};
double operator-(myMoney objX, myMoney objY)
{
	return objX.getMoney() - objY.getMoney();
}
myMoney operator--(myMoney &obj)
{
	if (obj.money > 1000)
	{
		obj.money -= 1000;
	}
	else
	{
		obj.money = 0;
	}
	return obj;
}
myMoney operator--(myMoney &obj, int)
{
	if (obj.time > 0)
	{
		obj.time--;
	}
	else
	{
		obj.time = 0;
	}
	return obj;
}

int main()
{
	setlocale(LC_ALL, "Russian");
	myMoney objA("Кот матроскин", 1200, 7, 1);
	objA.showAll();
	objA--;
	objA.showAll();
	objA--;
	objA.showAll();
	objA++;
	objA.showAll();
	--objA;
	objA.showAll();
	--objA;
	objA.showAll();
	++objA;
	objA.showAll();
	myMoney objB("Шарик", 1100, 8, 5);
	objB.showAll();
	myMoney objC;
	objC = objA + objB;
	objC.showAll();
	cout << "Разница в доходах: " << objC - objB << endl;

	system("pause>nul");
	return 0;
}

