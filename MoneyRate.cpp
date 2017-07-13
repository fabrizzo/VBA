// MassiveAndPointer.cpp: определ€ет точку входа дл€ консольного приложени€.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
using namespace std;

unsigned long getMoney(unsigned long m, double r)
{
	return m*(1 + r / 100);
}
unsigned long getMoney(unsigned long m, double r, int y)
{
	double s = m;
	for (int k = 1; k <= y;k++)
	{
		s *= (1 + r / 100);
	}
	return s;
}
unsigned long getMoney(unsigned long m, double r, int y, int n)
{
	return getMoney(m, r / n, y*n);
}
int main()
{
	setlocale(LC_ALL, "Russian");
	unsigned long money;
	double rate;
	cout << "¬ведите сумму вклада :";
	cin >> money;
	cout << "¬ведите процентную ставку : ";
	cin >> rate;
	cout << "Ќачальна€ сумма: " << money << endl;
	cout << "√одова€ ставка: " << rate << endl;
	cout << "¬клад на один год: " << getMoney(money,rate) << endl;
	cout << "¬клад на один год(начисл€етс€ 4 раза в год): " << getMoney(money, rate, 7, 4) << endl;
	cout << "¬клад на 3 лет: " << getMoney(money, rate, 3) << endl;
	cout << "¬клад на 3 лет (начислени€ 4 раза в год): " << getMoney(money, rate, 3, 4) << endl;
	cout << "¬клад на 5 лет: " << getMoney(money, rate, 5) << endl;
	cout << "¬клад на 5 лет (начислени€ 4 раза в год): " << getMoney(money, rate, 5, 4) << endl;
	cout << "¬клад на 7 лет: " << getMoney(money,rate,7) << endl;
	cout << "¬клад на 7 лет (начислени€ 4 раза в год): " << getMoney(money,rate,7,4) << endl;
	cout << "¬клад на 10 лет: " << getMoney(money, rate, 10) << endl;
	cout << "¬клад на 10 лет (начислени€ 4 раза в год): " << getMoney(money, rate, 10, 4) << endl;
	system("pause>nul");
	return 0;
}

