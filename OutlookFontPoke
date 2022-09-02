#!/bin/sh
#set -x

TOOL_NAME="Microsoft Outlook 365/2021/2019/2016 for Mac Default Font Changer"
TOOL_VERSION="2.1"

## Copyright (c) 2022 Microsoft Corp. All rights reserved.
## Scripts are not supported under any Microsoft standard support program or service. The scripts are provided AS IS without warranty of any kind.
## Microsoft disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a 
## particular purpose. The entire risk arising out of the use or performance of the scripts and documentation remains with you. In no event shall
## Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
## (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary 
## loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility
## of such damages.
## Feedback: pbowden@microsoft.com

# Constants
REGISTRY="$HOME/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB.reg"
OUTLOOKPATH="/Applications/Microsoft Outlook.app"
SCRIPTPATH=$( cd $(dirname $0) ; pwd -P )

function ShowUsage {
# Shows tool usage and parameters
	echo $TOOL_NAME - $TOOL_VERSION
	echo "Purpose: Sets the default compose and reply/forward fonts in Outlook 365/2021/2019/2016 for Mac"
	echo "Usage: OutlookFontPoke [-dump] <font-name> <font-size> <font-color>"
	echo "Example: OutlookFontPoke 'Helvetica' '11.0pt' 'gray'"
	echo
	exit 0
}

function CheckRegistryExists {
# Check if Registry exists
if [ ! -f "$REGISTRY" ]; then
	echo "WARNING: Registry DOES NOT exist at path $REGISTRY. Attempting to create..."
	mkdir "$HOME/Library/Group Containers/UBF8T346G9.Office"
	cp "$SCRIPTPATH/TemplateRegDB.reg" "$REGISTRY"
	if [ "$?" != "0" ]; then
		echo "ERROR: Registry could not be created."
		exit 1
	fi
fi
}

function CheckLaunchState {
# Checks to see if a process is running
	local RUNNING_RESULT=$(ps ax | grep -v grep | grep "$1")
	if [ "${#RUNNING_RESULT}" -gt 0 ]; then
		echo "1"
	else
		echo "0"
	fi
}

function GetNodeId {
# Get node_id value from Registry
	local NAME="$1"
	local PARENT="$2"
	local NODEVALUE=$(sqlite3 "$REGISTRY" "SELECT node_id from HKEY_CURRENT_USER WHERE name='$NAME' AND parent_id=$PARENT;")
	if [ "$NODEVALUE" == '' ]; then
		echo "0"
	else
		echo "$NODEVALUE"
	fi
}

function GetNodeVal {
# Get node value from Registry
	local NAME="$1"
	local NODEID="$2"
	local NODEVALUE=$(sqlite3 "$REGISTRY" "SELECT node_id from HKEY_CURRENT_USER_values WHERE name='$NAME' AND parent_id=$NODEID;")
	if [ "$NODEVALUE" == '' ]; then
		echo "0"
	else
		echo "$NODEVALUE"
	fi
}

function InsertNode {
# Insert new node into Registry
	local NAME="$1"
	local PARENT="$2"
	sqlite3 "$REGISTRY" "INSERT INTO HKEY_CURRENT_USER ('parent_id','name') VALUES ($PARENT,'$NAME');"
}

function InsertValue {
# Insert new value into Registry
	local NODE="$1"
	local NAME="$2"
	local TYPE="$3"
	local VALUE="$4"
	sqlite3 "$REGISTRY" "INSERT or REPLACE INTO HKEY_CURRENT_USER_values ('node_id','name','type','value') VALUES ($NODE,'$NAME',$TYPE,'$VALUE');"
}

function DeleteValue {
# Delete value from Registry
	local NAME="$1"
	local NODEID="$2"
	sqlite3 "$REGISTRY" "DELETE FROM HKEY_CURRENT_USER_values WHERE name='$NAME' and node_id=$NODEID;"
}

function GetValue {
# Get value from Registry
	local NAME="$1"
	local NODEID="$2"
	local NODEVALUE=$(sqlite3 "$REGISTRY" "SELECT value from HKEY_CURRENT_USER_values WHERE name='$NAME' AND node_id=$NODEID;")
	if [ "$NODEVALUE" == '' ]; then
		echo "0"
	else
		echo "$NODEVALUE"
	fi
}

# Evaluate command-line arguments
if [[ $# = 0 ]]; then
	ShowUsage
elif [[ $# = 1 ]]; then
	if [ "$1" == "-dump" ]; then
		DUMP=true
	else
		ShowUsage
	fi
else
	DUMP=false
	FONTNAME="$1"
	FONTSIZE="$2"
	FONTCOLOR="$3"
	
	if [ "$FONTCOLOR" == '' ]; then
		FONTCOLOR="windowtext"
	fi
fi

## Main
# Check that MicrosoftRegistryDB.reg actually exists. If it doesn't, attempt to create it.
CheckRegistryExists
# Walk the registry to find the id of the node that we need
KEY_SOFTWARE=$(GetNodeId "Software" '-1')
KEY_MICROSOFT=$(GetNodeId "Microsoft" "$KEY_SOFTWARE")
KEY_OFFICE=$(GetNodeId "Office" "$KEY_MICROSOFT")
KEY_VERSION=$(GetNodeId "16.0" "$KEY_OFFICE")
KEY_COMMON=$(GetNodeId "Common" "$KEY_VERSION")
KEY_MAILSETTINGS=$(GetNodeId "MailSettings" "$KEY_COMMON")
# The MailSettings node doesn't exist by default, so if it's not already there, create it
if [ "$KEY_MAILSETTINGS" == "0" ] && [ $DUMP = false ]; then
	InsertNode "MailSettings" "$KEY_COMMON"
fi

KEY_MAILSETTINGS=$(GetNodeId "MailSettings" "$KEY_COMMON")

# If the fonts are already set, remove the existing values
KEY_COMPOSEFONTCOMPLEX=$(GetValue "ComposeFontComplex" "$KEY_MAILSETTINGS")
if [ "$KEY_COMPOSEFONTCOMPLEX" != "0" ]; then
	if [ $DUMP = true ]; then
		echo "==Begin ComposeFontComplex=="
		echo "$KEY_COMPOSEFONTCOMPLEX"
		echo "==End ComposeFontComplex=="
	else
		DeleteValue "ComposeFontComplex" "$KEY_MAILSETTINGS"
		DeleteValue "ComposeFontSimple" "$KEY_MAILSETTINGS"
	fi
fi
KEY_REPLYFONTCOMPLEX=$(GetValue "ReplyFontComplex" "$KEY_MAILSETTINGS")
if [ "$KEY_REPLYFONTCOMPLEX" != "0" ]; then
	if [ $DUMP = true ]; then
		echo "==Begin ReplyFontComplex=="
		echo "$KEY_REPLYFONTCOMPLEX"
		echo "==End ReplyFontComplex=="
	else
		DeleteValue "ReplyFontComplex" "$KEY_MAILSETTINGS"
		DeleteValue "ReplyFontSimple" "$KEY_MAILSETTINGS"
	fi
fi

if [ $DUMP = false ]; then
	# Set new font values - first one is for the Compose Font, the second is for the Reply/Forward font
	InsertValue "$KEY_MAILSETTINGS" "ComposeFontComplex" "3" "<html><head><style>/* Style Definitions */span.PersonalComposeStyle{mso-style-name:\"Personal Compose Style\";mso-style-type:personal-compose;mso-style-noshow:yes;mso-style-unhide:no;mso-ansi-font-size:$FONTSIZE;mso-bidi-font-size:11.0pt;font-family:\"$FONTNAME\";mso-ascii-font-family:\"$FONTNAME\";mso-hansi-font-family:\"$FONTNAME\";mso-bidi-font-family:\"Times New Roman\";mso-bidi-theme-font:minor-bidi;color:$FONTCOLOR;font-weight:normal;font-style:normal;text-decoration:none;text-underline:none;}--></style></head></html>"
	InsertValue "$KEY_MAILSETTINGS" "ReplyFontComplex" "3" "<html><head><style>/* Style Definitions */span.PersonalReplyStyle{mso-style-name:\"Personal Reply Style\";mso-style-type:personal-reply;mso-style-noshow:yes;mso-style-unhide:no;mso-ansi-font-size:$FONTSIZE;mso-bidi-font-size:11.0pt;font-family:\"$FONTNAME\";mso-ascii-font-family:\"$FONTNAME\";mso-hansi-font-family:\"$FONTNAME\";mso-bidi-font-family:\"Times New Roman\";mso-bidi-theme-font:minor-bidi;color:$FONTCOLOR;font-weight:normal;font-style:normal;text-decoration:none;text-underline:none;}--></style></head></html>"

	echo "Default Outlook font set successfully."

	# If Outlook is already running, show a warning that the settings won't take effect until a restart occurs
	RUNSTATE=$(CheckLaunchState "$OUTLOOKPATH")
	if [ "$RUNSTATE" == "1" ]; then
		echo "Outlook must be restarted to read the new font preference."
	fi
fi


exit 0
