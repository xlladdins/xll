// xll.t.cpp - test xll add-in library

extern "C"
int __declspec(dllexport) __stdcall
xlAutoOpen(void);

int main()
{
	xlAutoOpen();

	return 0;
}