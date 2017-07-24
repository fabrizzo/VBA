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

	myMoney(string n, double m, double r, int t)
	{
		setAll(n, m, r, t);
	}
	myMoney()
	{
		setAll("", 0, 0, 0);
	}

};
class BigMoney : public myMoney
{
public:
	int periods;
	double getMoney()
	{
		double s = money;
		for (int k = 1;k <= time*periods;k++)
		{
			s *= (1 + rate / 100 / periods);
		}
		return s;
	}
	void showAll()
	{
		cout << "Имя: " << name << endl;
		cout << "Вклад: " << money << endl;
		cout << "Ставка (%): " << rate << endl;
		cout << "Период (лет): " << time << endl;
		cout << "Начислений в год: " << periods<< endl;
		cout << "Итоговая сумма: " << getMoney() << endl;
		for (int k = 1;k <= 35;k++)
		{
			cout << "-";
		}
		cout << endl;
	}
	void setAll(string n, double m, double r, int t, int p)
	{
		myMoney::setAll(n, m, r, t);
		periods = p;
	}
	BigMoney(string n, double m, double r, int t, int p = 1) : myMoney(n, m, r, t)
	{
		periods = p;
	}
	BigMoney(): myMoney() 
	{
	periods = 1;
	}
};


int main()
{
	setlocale(LC_ALL, "Russian");
	myMoney objA("Кот матроскин", 1200, 8, 5);
	BigMoney objB("Дядя Фёдор", 1000, 7, 6, 2);
	BigMoney objC("Пес Шарик", 1500, 6, 8);
	BigMoney objD;
	objD.setAll("Почтальон Печкин", 800, 10, 3, 4);
	objA.showAll();
	cout << endl;
	objB.showAll();
	cout << endl;
	objC.showAll();
	cout << endl;
	objD.showAll();
	cout << endl;



	system("pause>nul");
	return 0;
}

