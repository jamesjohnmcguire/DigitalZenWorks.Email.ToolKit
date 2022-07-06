#pragma once

namespace MapiLibrary
{
	#if defined _WIN32 || defined __CYGWIN__
		#ifdef MAPILIBRARY_EXPORTS
			#if defined WIN32
				#define LIB_API(RetType) extern "C" __declspec(dllexport) RetType
				#define DllExport __declspec(dllexport)
			#else
				#define LIB_API(RetType) extern "C" RetType __attribute__((visibility("default")))
			#endif

			#ifdef __GNUC__
				#define EXPORT_API extern "C" __attribute__ ((dllexport))
			#else
				// Note: actually gcc seems to also supports this syntax.
				#define EXPORT_API extern "C" __declspec (dllexport)
			#endif
		#else
			#define DllExport

			#if defined WIN32
				#define LIB_API(RetType) extern "C" __declspec(dllimport) RetType
			#else
				#define LIB_API(RetType) extern "C" RetType
			#endif

			#ifdef __GNUC__
				#define EXPORT_API extern "C" __attribute__ ((dllimport))
			#else
				#define EXPORT_API extern "C" __declspec (dllimport)
			#endif
		#endif
	#else
		#if __GNUC__ >= 4
			#define EXPORT_API extern "C" __attribute__ ((visibility ("default")))
			#define LIB_API(RetType) RetType
		#else
			#define EXPORT_API
			#define LIB_API(RetType) RetType
		#endif
	#endif
}
