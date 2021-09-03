#include<iostream>
#include<cmath>
#include<vector>
typedef short int digit;
std::vector<digit> getNumber(std::vector<digit>& v) {
tryAgain:
	std::string str;
	std::cout << "input number";
	std::cin >> str;
	v.resize((str.length()));
	for (int i = 0; i < str.length(); i++) {
		if (str[i] < '0' || str[i] > '9') {
			std::cout << "input correct number";
			goto tryAgain;
			break;
		}
		v.at(i) = str[i] - '0';
	}
	std::cin.ignore();
	return v;
}

char getOperand() {
	char operand;
	while (true) {
		std::cout << "input operation ";
		std::cin >> operand;
		std::cin.ignore(32767, '\n');
		if (operand == '+' || operand == '/' || operand == '-' || operand == '*' || operand == '**' || operand == '&') {
			return operand;
		}
		std::cout << "input correct operation symbol " << std::endl;
	}
	return operand;
}
std::vector<digit> minus(std::vector<digit> a, std::vector<digit> b) {
	int flag = 0;
	int zflag = 0;
	int j;
	std::vector<digit> result;
	if (a.size() > b.size()) {
		for (int i = a.size() - b.size(); i > 0; --i) {
			b.insert(b.begin(), 0);
		}
	}
	else if (b.size() > a.size()) {
		for (int i = b.size() - a.size(); i > 0; --i) {
			a.insert(a.begin(), 0);
		}
	}
	for (int i = 0; i < a.size(); i++) {
		if (a[i] > b[i]) {
			zflag = 0;
			break;
		}
		else if (b[i] > a[i]) {
			zflag = 1;
			break;
		}
	}
	if (zflag == 0) {
		for (int i = a.size() - 1; i >= 0; i--) {
			if (a[i] - b[i] >= 0) {
				result.insert(result.begin(), a[i] - b[i]);
			}
			else if (a[i] - b[i] < 0) {
				result.insert(result.begin(), (a[i] + 10) - b[i]);
				j = i - 1;
				if (a[j] > 0) {
					a.at(j) = a.at(j) - 1;
				}
				else if (a[j] == 0) {
					while (a[j] == 0) {
						a[j] = 9;
						j--;
					}
					a[j] = a[j] - 1;
				}
			}
		}
	}
	if (zflag == 1) {
		for (int i = a.size() - 1; i >= 0; i--) {
			if (b[i] - a[i] >= 0) {
				result.insert(result.begin(), b[i] - a[i]);
			}
			else if (b[i] - a[i] < 0) {
				result.insert(result.begin(), (b[i] + 10) - a[i]);
				j = i - 1;
				if (b[j] > 0) {
					b[j] = b[j] - 1;
				}
				else if (b[j] == 0) {
					while (b[j] == 0) {
						b.at(j) = 9;
						j--;
					}
					b[j] = b[j] - 1;
				}
			}
		}
		std::cout << '-';
	}
	int l = 0;
	while (result.at(l) == 0) {
		if (result.at(l) == '-') {
			l++;
			result.erase(result.begin() + l);
		}
		result.erase(result.begin());
	}

	return result;
	result.clear();
	result.shrink_to_fit();
}
std::vector<digit> sum(std::vector<digit> a, std::vector<digit> b) {
	int flag = 0;
	std::vector<digit> result;
	if (a.size() > b.size()) {
		for (int i = a.size() - b.size(); i > 0; i--) {
			b.insert(b.begin(), 0);
		}
	}
	else if (b.size() > a.size()) {
		for (int i = b.size() - a.size(); i > 0; i--) {
			a.insert(a.begin(), 0);
		}
	}
	for (int i = a.size() - 1; i >= 0; i--) {
		result.insert(result.begin(), (a[i] + b[i] + flag) % 10);
		if (i == 0 && a[i] + b[i] + flag > 9)
			result.insert(result.begin(), 1);
		if ((a[i] + b[i] + flag) > 9)
			flag = 1;
		else
			flag = 0;
	}
	return result;
	result.clear();
	result.shrink_to_fit();
}
void calc(std::vector<digit> a, char operand, std::vector<digit> b) {
	std::vector<digit> res;

	if (operand == '+') {
		res = sum(a, b);
		for (auto const& element : res)
			std::cout << element << ' ';
	}
	if (operand == '-') {
		res = minus(a, b);
		for (auto const& element : res)
			std::cout << element << ' ';
	}
}

int main() {
	std::vector<digit> x;
	std::vector<digit> y;
	getNumber(x);
	getNumber(y);
	calc(x, getOperand(), y);
	system("pause");
}