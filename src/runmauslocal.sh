#!/bin/bash

export LC_ALL="C"
MAUS="$(dirname "$0")/../external/maus/maus"

echoerr() { echo "$(date +"####[%y-%m-%d %T]####")" "$@" 1>&2; }

echocmd() {
	echoerr "**** exec:" "$@"
	"$@"
}

function callmaus() {
	prefix=${1%%.*}
	if [ ! -e "$prefix.par" ]; then
	    echo "No transcription found for $prefix.wav. Skipping..."
	    return 0
	fi
    if ! "$MAUS" v=0 OUT="$prefix.par" OUTFORMAT=mau-append SIGNAL="$1" BPF="$prefix.par" \
        USETRN=no NOINITIALFINALSILENCE="$2" LANGUAGE="$3" >"$prefix.log" 2>&1; then
        cat "$prefix.log"
        echoerr "Error/Warning processing $prefix.wav"
        return 0
    fi
}

tmpdir="$1"
case "$2" in
    True) nosil="yes" ;;
    False) nosil="no" ;;
    *) echo "Error: Unknown boolean value \"$2\"" 1>&2; exit 1 ;;
esac

for fwav in $tmpdir/iv*.wav; do
    callmaus "$fwav" "$nosil" "$3"
done
