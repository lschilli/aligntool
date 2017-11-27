form params
    sentence infilename
    natural beepchannel 2
    sentence refbeepname
    real silencethreshold -25
    real minsoundingduration 0.18
    boolean seekflank 1
endform
refbeep = Read from file: refbeepname$
beeplength = Get total duration
beepsampling = Get sampling frequency
infile = Read from file: infilename$
beepchannel = Extract one channel: beepchannel
removeObject: infile
segments = To TextGrid (silences): 400, 0.01, silencethreshold, 0.03, minsoundingduration, "speech", "beep"
select 'beepchannel'
beepchansampling = Get sampling frequency
# different sampling freqs even per language: resample if necessary
if beepsampling <> beepchansampling
    beepchan2 = Resample: beepsampling, 1
    removeObject: beepchannel
    beepchannel = beepchan2
endif
select 'refbeep'
plus 'beepchannel'
# Not needed for a clean signal but helps to identify the peek
corrsig = Cross-correlate... "peak 0.99" similar
select 'segments'
numIntervals = Get number of intervals: 1
for i from 1 to numIntervals
    select 'segments'
    text$ = Get label of interval: 1, i
    if text$ = "beep"
        startTime = Get start point: 1, i
        endTime = Get end point: 1, i
        select 'corrsig'
        # The maximum of the correlation is typically at the beginning of the intervall, we search around there
        maxT = Get time of maximum: (startTime-beeplength*0.5), (startTime+beeplength*0.5), "Parabolic"
        maxV = Get value at time: 1, maxT, "Nearest"
        # Refine start time inside a 10ms window. threshold based
        select 'beepchannel'
        sampPeriod = Get sampling period
        if seekflank = 1
            x = maxT-0.01
            xmax = maxT+0.01
            extrVal = Get absolute extremum: x, xmax, "Parabolic"
            thresh = extrVal * 0.5
            while x < xmax
                val = Get value at time: 1, x, "Cubic"
                if abs(val) > thresh
                    maxT = x
                    goto ready
                endif
                x = x + sampPeriod
            endwhile
            label ready
        endif
        writeInfoLine: maxT, tab$, maxT+beeplength, tab$, maxV
    endif
endfor

