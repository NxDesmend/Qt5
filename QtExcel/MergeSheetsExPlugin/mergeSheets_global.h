#pragma once

#include <QtCore/qglobal.h>

#ifndef BUILD_STATIC
# if defined(MERGESHEETS_LIB)
#  define MERGESHEETS_EXPORT Q_DECL_EXPORT
# else
#  define MERGESHEETS_EXPORT Q_DECL_IMPORT
# endif
#else
# define MERGESHEETS_EXPORT
#endif
