// XOR.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <string>

using namespace std;

string XOR(string data, char key[])
{
	string xorstring = data;
	for (int i = 0; i < xorstring.size();i++)
	{
		xorstring[i] = data[i] ^ key[i % (sizeof(key) / sizeof(char))];
	}
	return xorstring;
}

int main()
{
	string dexorstring;
	setlocale(LC_ALL, "Russian");
	char key[3] = {'K', 'E', 'Y' };
	cout <<"Защифрованно: "<< XOR("Hacker", key) << endl;
	dexorstring = XOR("Hacker", key);
	dexorstring = XOR(dexorstring, key);
	cout << "Расшифрованно: " << dexorstring << endl;
	getchar();
	return 0;
}

