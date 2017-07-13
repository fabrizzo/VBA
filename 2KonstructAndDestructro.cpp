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
	myMoney()
	{
		name = "Кот матроскин";
		money = 100;
		rate = 5;
		time = 1;
		cout << "Cоздан новый обьект:\n";
		showAll();
	}
	myMoney(string n, double m, double r, int t)
	{
		setAll(n, m, r, t);
		cout << "Cоздан новый обьект:\n";
		showAll();
	}
	~myMoney()
	{
		cout << "Обьект для \"" << name << "\" удален\n";
		for (int k = 1; k <= 35;k++)
		{
			cout << "*";
		}
		cout << endl;
	}

};
void postman()
{
	myMoney objD("Почтальон Печкин", 200, 3, 2);
}
int main()
{
	setlocale(LC_ALL, "Russian");
	myMoney objA;
	myMoney objB("Дядя Федор",1500,8,7);
	postman();
	myMoney* objC = new myMoney("Пес шарик", 1200, 6, 9);
	cout << "Все обьекты созданы\n";
	delete objC;
	cout << "Выполнение программы завершено\n";
	cout << endl;


	system("pause>nul");
	return 0;
}

