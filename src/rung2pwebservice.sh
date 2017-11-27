#!/bin/bash

function rung2p() {
	infile="$1"
	prefix=${1%%.*}
	language="$2"

	response=$(curl -sS -X POST -H 'content-type: multipart/form-data' \
	-F com=yes -F align=no -F stress=no -F lng="$language" -F syl=no -F embed=no \
	-F iform=txt -F i=@"$infile" -F tgrate=16000 -F nrm=no -F oform=tab \
	-F map=no -F featset=standard -F tgitem=ort \
	'https://clarin.phonetik.uni-muenchen.de/BASWebServices/services/runG2P')

	#echo "$response"
	downloadlink=$(xmllint --xpath '//downloadLink/text()' <(echo "$response"))
	if [ -z "$downloadlink" ]; then
	    echo "ERROR"
	    echo "$response"
	    exit 1
	fi
	if curl -sS "$downloadlink" > "$prefix.tab"; then
        echo "GOT: $downloadlink"
    fi
}

rung2p "$@"
