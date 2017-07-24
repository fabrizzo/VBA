#include "stdafx.h"
#include <iostream>
#include <string>
#include <cstdlib>
using namespace std;
class MyClass {
public:
	char code;
	MyClass* next;
	MyClass(MyClass &obj) {
		obj.next = this;
		code = obj.code + 1;
	}
	MyClass(char s) {
		code = s;
	}
	~MyClass() {
		if (next) {
			delete next;
		}
		cout << "Объект с полем " << code << " удален \n";
	}
	void show() {
		cout << code << " ";
		if (next) {
			next->show();
		}
	}
};
using namespace std;

int main()
{
	setlocale(LC_ALL, "Russian");
	int n = 10;
	MyClass* pnt = new MyClass('!');
	MyClass *p;
	p = pnt;
	for (int k = 1;k <= n;k++) {
		p = new MyClass(*p);
	}
	p->next = 0;
	pnt->show();
	cout << endl;
	delete pnt;

	system("pause>nul");
	return 0;
}