// MassiveAndPointer.cpp: ���������� ����� ����� ��� ����������� ����������.
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
		cout << "���: " << name << endl;
		cout << "�����: " << money << endl;
		cout << "������ (%): " << rate << endl;
		cout << "������ (���): " << time << endl;
		cout << "�������� �����: " << getMoney() << endl;
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
		name = "��� ���������";
		money = 100;
		rate = 5;
		time = 1;
		cout << "C����� ����� ������:\n";
		showAll();
	}
	myMoney(string n, double m, double r, int t)
	{
		setAll(n, m, r, t);
		cout << "C����� ����� ������:\n";
		showAll();
	}
	~myMoney()
	{
		cout << "������ ��� \"" << name << "\" ������\n";
		for (int k = 1; k <= 35;k++)
		{
			cout << "*";
		}
		cout << endl;
	}

};
void postman()
{
	myMoney objD("��������� ������", 200, 3, 2);
}
int main()
{
	setlocale(LC_ALL, "Russian");
	myMoney objA;
	myMoney objB("���� �����",1500,8,7);
	postman();
	myMoney* objC = new myMoney("��� �����", 1200, 6, 9);
	cout << "��� ������� �������\n";
	delete objC;
	cout << "���������� ��������� ���������\n";
	cout << endl;


	system("pause>nul");
	return 0;
}

