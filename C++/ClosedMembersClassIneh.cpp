// ClosedMembersClassIneh.cpp: определяет точку входа для консольного приложения.
//

#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;


class Alpha
{
private: 
	char symb;
public:
	Alpha(char s) {
		symb = s;
	}
	void show() {
		cout << "Класс Alpha: " << symb << endl;
	}
	void set(char s) {
		symb = s;
	}
};

class Bravo : public Alpha {
public:
	int num;
	Bravo(char s, int n ) :Alpha(s) {
		num = n;
	}
	void showAll() {
		show();
		cout << "Класс Bravo: " << num << endl;
	}
};

int main()
{
	setlocale(LC_ALL, "Russian");
	Bravo obj('A', 100);
	obj.showAll();
	obj.set('Z');
	obj.showAll();
	system("pause>nul");
	return 0;
}

