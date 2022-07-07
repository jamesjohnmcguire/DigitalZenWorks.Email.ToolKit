#pragma once

namespace MapiLibrary
{
	#ifdef MAPILIBRARY_EXPORTS
		#define API extern "C" __declspec(dllexport)
		#define DllExport __declspec(dllexport)
	#else
		#define API extern "C" __declspec(dllimport)
		#define DllExport
	#endif
}
