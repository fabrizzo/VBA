// VirtualMethod.cpp: определяет точку входа для консольного приложения.
//

#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;
class Alpha {
public:
	virtual void show() {
		cout << "Класс Alpha" << endl;
	}
	void showAll() {
		show();
	}
};
class Bravo : public Alpha {
public:
	void show() {
		cout << "Класс Bravo" << endl;
	}
};
int main()
{
	setlocale(LC_ALL, "Russian");
	Bravo obj;
	obj.show();
	obj.showAll();
	system("pause>num");
    return 0;
}

