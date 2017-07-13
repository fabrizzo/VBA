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
private:
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
public:
	void showAll()
	{
		cout << "Имя: " << name << endl;
		cout << "Вклад: " << money << endl;
		cout << "Ставка (%): " << rate << endl;
		cout << "Период (лет): " << time << endl;
		cout << "Итоговая сумма: " << getMoney() << endl;

	}
	void setAll(string n, double m, double r, int t)
	{
		name = n;
		money = m;
		rate = r;
		time = t;
	}
};
int main()
{
	setlocale(LC_ALL, "Russian");
	myMoney objA, objB;
	objA.setAll("Иванов Иван Иванович", 1000, 8, 5);
	objB.setAll("Петров Петр Петрович", 1200, 7, 4);
	objA.showAll();
	cout << endl;
	objB.showAll();

	system("pause>nul");
	return 0;
}

