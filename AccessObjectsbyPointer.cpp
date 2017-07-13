// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <string>

using namespace std;
class MyClass {
public:
	string name;
	int number;
	void show() {
		cout << "Поле Имя:" << name << endl;
		cout << "Поле Номер:" << number << endl;
		for (int k = 0;k < 35;k++)
		{
			cout << "-";
		}
		cout << "\n";
	}
};

int main()
{
	setlocale(LC_ALL, "Russian");
	MyClass objA, objB;
	MyClass* p;
	p = &objA;
	p->name = "Дмитрий";
	p->number = 111;
	p->show();
	p = &objB;
	p->name = "Ольга";
	p->number = 222;
	p->show();
	cout << "Проверяем объекты\n";
	objA.show();
	objB.show();
	system("pause>nul");
	return 0;
}

