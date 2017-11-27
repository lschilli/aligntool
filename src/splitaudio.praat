form params
    sentence wavfile
    sentence textgridfile
    sentence outfileprefix tmp/iv
    natural speechchannel 1
    sentence speechtier seg.speech
    boolean noisefilter 0
endform

inwav = Read from file: wavfile$
speechchannel = Extract one channel: speechchannel
if noisefilter = 1
    speechchannel = Remove noise: 0, 3, 0.025, 80, 8000, 40, "Spectral subtraction"
endif
intextgrid = Read from file: textgridfile$

selectObject(intextgrid)

.n = Get number of tiers
speechtieridx = -1
for .i to .n
  .name$ = Get tier name: .i
  if .name$ == speechtier$
    speechtieridx = .i
    .i += .n
  endif
endfor

if speechtieridx == -1
	exitScript: "Unable to find a tier named ", speechtier$, " in ", textgridfile$
endif

selectObject(speechchannel)
plusObject(intextgrid)
Extract intervals where: speechtieridx, "yes", "is equal to", "speech"

chunk_count = numberOfSelected()
for i to chunk_count
	chunks[i]  = selected(i)
endfor

for i from 1 to chunk_count
	selectObject(chunks[i])
	startTime = Get start time
	endTime = Get end time
	writeInfoLine: string$(startTime) + tab$ + string$(endTime) + tab$ + outfileprefix$ + string$(i) + ".wav"
	Save as WAV file: outfileprefix$ + string$(i) + ".wav"
endfor
