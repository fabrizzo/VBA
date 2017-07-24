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

	}
};
int main()
{
	setlocale(LC_ALL, "Russian");
	myMoney obj;
	obj.name = "Иванов Иван Иванович";
	obj.money = 1000;
	obj.rate = 8;
	obj.time = 5;
	obj.showAll();
	system("pause>nul");
	return 0;
}

