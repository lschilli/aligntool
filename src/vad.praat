form params
    sentence wavfile
    natural speechchannel 1
    boolean noisefilter 0
    positive trainbegin 1
    positive trainwindow 1
    real threshold 30
    real minSilenceDuration 0.02
    real minSoundingDuration 0.02
endform

inwav = Read from file: wavfile$
speechchan = Extract one channel: speechchannel
removeObject: inwav
if noisefilter = 1
    speechchan2 = Remove noise: trainbegin, trainbegin+trainwindow, 0.01, 80, 16000, 10, "Spectral subtraction"
    removeObject: speechchan
    speechchan = speechchan2
endif

intensityObj = To Intensity: 100, 0.01, "no"
allintensitymax = Get maximum: 0, 0, "Parabolic"
sndb = max(allintensitymax-threshold, 0.01)
writeInfoLine: "silence threshold: " + string$(-sndb) + " db"
vadtier = To TextGrid (silences): -sndb, minSilenceDuration, minSoundingDuration, "", "speech"

selectObject(speechchan)
plusObject(vadtier)
Extract intervals where: 1, "yes", "is equal to", "speech"

chunk_count = numberOfSelected()
for i to chunk_count
	chunks[i]  = selected(i)
endfor

for i from 1 to chunk_count
	selectObject(chunks[i])
	startTime = Get start time
	endTime = Get end time
	selectObject(intensityObj)
    meandb = Get mean: startTime, endTime, "dB"
	writeInfoLine: "chunk", tab$, startTime, tab$, endTime, tab$, meandb
endfor

selectObject: intensityObj
framecnt = Get number of frames
for i from 1 to framecnt
    t = Get time from frame number: i
    v = Get value in frame: i
    writeInfoLine:  "itn", tab$, t, tab$, v
endfor
