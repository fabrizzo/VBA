// Generalized functions.cpp: определяет точку входа для консольного приложения.
//

#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;

template<class Z, class Y> void show(Z z, Y y) {
	cout << "Функция с 2 аргументами\n";
	cout << "Значение аргумента 1: " << z << endl;
	cout << "Значение аргумента 2: " << y << endl;
}


template<class X>void show(X x) {
	cout << "Функция с одним аргументом\n";
	cout << "Значение аргумента: " << x << endl;
}

int main()
{
	setlocale(LC_ALL, "Russian");
	show('A');
	show(123);
	show("Текст");
	show(321, "Текст");
	show('B', 456);
	show('C', 'D');

	



	system("pause>nul");
	return 0;
}
