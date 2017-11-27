#!/bin/bash

function callmaus() {
	for i in 1 2 3; do
	prefix=${1%%.*}
	if [ ! -e "$prefix.par" ]; then
	    echo "No transcription found for $prefix.wav. Skipping..."
	    break
	fi
    response=$(curl -sS -X POST -H 'content-type: multipart/form-data' \
    -F OUTIPA=false \
    -F NOINITIALFINALSILENCE="$2" \
    -F INFORMAT=bpf \
    -F LANGUAGE="$3" \
    -F OUTSYMBOL=sampa \
    -F MINPAUSLEN=5 \
    -F USETRN=false \
    -F SIGNAL=@"$prefix.wav" \
    -F STARTWORD=0 \
    -F ENDWORD=999999 \
    -F INSPROB=0.0 \
    -F OUTFORMAT=mau-append \
    -F BPF=@"$prefix.par" \
    -F INSKANTEXTGRID=false \
    -F WEIGHT=default \
    -F MAUSSHIFT=default \
    -F MODUS=standard \
    -F INSORTTEXTGRID=true \
    'https://clarin.phonetik.uni-muenchen.de/BASWebServices/services/runMAUS')

	downloadlink=$(xmllint --xpath '//downloadLink/text()' <(echo "$response"))
	if [ -z "$downloadlink" ]; then
	    echo "$response"
	    echo "ERROR, once again..."
	    continue
	fi
	if curl -sS "$downloadlink" > "$prefix.par"; then
        echo "GOT: $downloadlink"
        break
    fi
	done
}

tmpdir="$1"
case "$2" in
    True) nosil="true" ;;
    False) nosil="false" ;;
    *) echo "Error: Unknown boolean value \"$2\"" 1>&2; exit 1 ;;
esac

for fwav in $tmpdir/iv*.wav; do
    callmaus "$fwav" "$nosil" "$3" &
    while (( $(jobs -rp | wc -l) > 9 )); do sleep 0.5; done
done

wait
