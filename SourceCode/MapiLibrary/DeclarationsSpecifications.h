#pragma once

namespace MapiLibrary
{
	#ifdef MAPILIBRARY_EXPORTS
		#define API(RetType) extern "C" __declspec(dllexport) RetType
		#define DllExport __declspec(dllexport)

		#ifdef __GNUC__
			#define EXPORT_API extern "C" __attribute__ ((dllexport))
		#else
			// Note: actually gcc seems to also supports this syntax.
			#define EXPORT_API extern "C" __declspec (dllexport)
		#endif
	#else
		#define DllExport

		#define API(RetType) extern "C" __declspec(dllimport) RetType

		#ifdef __GNUC__
			#define EXPORT_API extern "C" __attribute__ ((dllimport))
		#else
			#define EXPORT_API extern "C" __declspec (dllimport)
		#endif
	#endif
}
