# -*- coding: utf-8 -*-
"""
/***************************************************************************
 SPREADSHEET
                                 A QGIS plugin
 This plugin will  scan all spreadsheet in specified folder for a particular circuit 
                             -------------------
        begin                : 2017-05-05
        copyright            : (C) 2017 by Rau.D
        email                : raul.dasgupta@vocus.com.au
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load SPREADSHEET class from file SPREADSHEET.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .spreadsheet import SPREADSHEET
    return SPREADSHEET(iface)
