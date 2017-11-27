#!/usr/bin/python3

import sys
import argparse
import logging
import os
import shutil
import subprocess
import tempfile
import csv
from collections import namedtuple, Counter
from collections import deque
sys.path.append(os.path.join(os.path.dirname(sys.argv[0]), "..", "lib", "python"))
import tgt
import util
import xlsbatch
import re
import gui


cachedir = "."
BeepTuple = namedtuple("beep_iv_tuple", ["t_start", "t_end", "correlation"])
IntensityVal = namedtuple("IntensityVal", ["t", "intensity"])


def segment_beeps(infile, outfile, wavfile, beepchannel, refbeep,
                  silencethreshold, minsoundingduration, seekflank, mincorrelation):
    logging.info("Segmenting beeps in %s" % wavfile)
    tmpdir = tempfile.mkdtemp()
    try:
        duration, _ = util.get_wav_duration(wavfile)
        tg, tier = util.init_textgrid(infile, duration, "seg.beep")
        beepsegmentscript = os.path.join(os.path.dirname(sys.argv[0]), "beepsegment.praat")
        beeplists = []
        for rb in refbeep:
            logging.info("Running correlation search for beep %s" % rb)
            result = util.call_check(["praat", "--run", beepsegmentscript, os.path.realpath(wavfile), str(beepchannel),
                                      os.path.realpath(rb), str(silencethreshold), str(minsoundingduration),
                                      str(int(seekflank))], True)
            beeplist = []
            for line in result.decode().split("\n"):
                if line:
                    t_start, t_end, correlation = line.split("\t")
                    bt = BeepTuple(float(t_start), float(t_end), float(correlation))
                    beeplist.append(bt)
            beeplists.append(beeplist)

        max_index_list = []
        for btlist in zip(*beeplists):
            bt = max(btlist, key=lambda t: t.correlation)
            max_index_list.append(btlist.index(bt))
            if bt.correlation < mincorrelation:
                logging.warning("low correlation (%.02f) with reference beep at %.4f seconds" %
                                (bt.correlation, bt.t_start))
                continue
            overlapping_beeps = tier.get_annotations_between_timepoints(bt.t_start, bt.t_end,
                                                                        left_overlap=True, right_overlap=True)
            can_add = True
            for x in overlapping_beeps:
                if x.correlation < bt.correlation:
                    tier.delete_annotation_by_start_time(x.start_time)
                else:
                    can_add = False
            if can_add:
                anno = tgt.Annotation(bt.t_start, bt.t_end, "beep")
                anno.correlation = bt.correlation
                tier.add_annotation(anno)
        cnt = Counter(max_index_list)
        logging.info("Found %d beep instances with (index, count) distribution %s" %
                     (len(max_index_list), sorted(cnt.items())))
        tier = tier.get_copy_with_gaps_filled(empty_string="speech")
        if len(tier) > 0:
            tier[0].text = ""
            logging.info("Note: setting first (silence) interval empty")
        tg.add_tier(tier)
        logging.info("Writing %s" % outfile)
        tgt.io.write_to_file(textgrid=tg, filename=outfile, format="long")
    finally:
        shutil.rmtree(tmpdir)


def segment_speech_praat(wavfile, channel=1, denoise=False, trainbegin=0, trainwindow=1, threshold=50,
                         min_sil_duration=0.02, min_snd_duration=0.02):
    vadscript = os.path.join(os.path.dirname(sys.argv[0]), "vad.praat")
    result = util.call_check(["praat", "--run", vadscript, os.path.realpath(wavfile), str(channel),
                              str(int(denoise)), str(trainbegin), str(trainwindow), str(threshold),
                              str(min_sil_duration), str(min_snd_duration)], True)
    speech_chunks = []
    intensities = []
    for line in result.decode().split("\n"):
        items = line.split("\t")
        if line.startswith("silence threshold"):
            logging.info(line)
        elif len(items) > 2:
            if items[0] == "chunk":
                iv = tgt.Interval(float(items[1]), float(items[2]), items[3])
                iv.as_db = float(items[3])
                speech_chunks.append(iv)
            elif items[0] == "itn":
                intensities.append(IntensityVal(float(items[1]), float(items[2])))
    return speech_chunks, intensities


def window(seq, n):
    it = iter(seq)
    win = deque((next(it, None) for _ in range(n)), maxlen=n)
    yield win
    append = win.append
    for e in it:
        append(e)
        yield win


def find_silence_level(intensities, window_size):
    min_max_val = None
    for w in window(intensities, int(window_size*100)):
        max_int = max(w, key=lambda x: x.intensity)
        if min_max_val is None or min_max_val.intensity > max_int.intensity:
            min_max_val = max_int
    return min_max_val.intensity


def filter_chunks(speech_chunks, silencelevel, speechthresh=0.8):
    dbvalues = [float(x.text) - silencelevel for x in speech_chunks]
    dbfilterthreshold = silencelevel + (sum(dbvalues) / len(dbvalues) * speechthresh)
    logging.info("speech filtering threshold: %s" % dbfilterthreshold)
    logging.info("vad segments: %s" % len(speech_chunks))

    dbfilteredivs = tgt.IntervalTier(name="seg.filt")
    for siv in [x for x in speech_chunks if float(x.text) > dbfilterthreshold]:
        dbfilteredivs.add_annotation(siv)

    logging.info("dropped %s intervals" % (len(speech_chunks)-len(dbfilteredivs)))
    return dbfilteredivs


def segment_speech(infile, outfile, wavfile, channel, filtertiername, shiftonset, shiftoffset, denoise,
                   trainbegin, trainwindow, speechthresh, snradd):
    logging.info("Segmenting speech in %s" % wavfile)
    duration, _ = util.get_wav_duration(wavfile)
    tg, tier = util.init_textgrid(infile, duration, "seg.speech")

    logging.info("Floor estimation...")
    _, intensities = segment_speech_praat(wavfile, channel,
                                          denoise=denoise, trainbegin=trainbegin, trainwindow=trainwindow)
    silencelevel = find_silence_level(intensities, trainwindow) + snradd
    logging.info("estimated floor noise level: %s" % silencelevel)
    logging.info("Segmentation...")
    speech_chunks, intensities = segment_speech_praat(wavfile, channel, threshold=silencelevel,
                                                      denoise=denoise, trainbegin=trainbegin, trainwindow=trainwindow)
    for iv in speech_chunks:
        tier.add_annotation(iv)

    dbvalues = [x.as_db - silencelevel for x in tier]
    dbfilterthreshold = silencelevel + (sum(dbvalues) / len(dbvalues) * speechthresh)
    logging.info("speech filtering threshold: %s" % dbfilterthreshold)
    logging.info("vad segments: %s" % len(tier))

    if filtertiername is None:
        filtertier = tgt.IntervalTier()
        filtertier.add_annotation(tgt.Annotation(tier.start_time, tier.end_time, "speech"))
        filtertiername = "<all>"
    else:
        filtertier = tg.get_tier_by_name(filtertiername)
    resulttier = tgt.IntervalTier(name="seg.speech")
    speechsegments = [s for s in filtertier if s.text == "speech"]
    logging.info("expected speech segments: %s" % len(speechsegments))
    stats_filtered = 0
    stats_all = 0
    for speechseg in speechsegments:
        speechivs = tier.get_annotations_between_timepoints(speechseg.start_time, speechseg.end_time)
        if len(speechivs) == 0:
            speechivs = tier.get_annotations_between_timepoints(speechseg.start_time, speechseg.end_time,
                                                                left_overlap=True, right_overlap=True)
            if len(speechivs) > 0:
                logging.warning("Speech segments overlap with the boundaries of %s in %s. "
                                "VAD problem? Shortening..." % (
                                    speechseg, filtertiername))
                for siv in speechivs:
                    siv.start_time = max(speechseg.start_time, siv.start_time)
                    siv.end_time = min(speechseg.end_time, siv.end_time)

        if len(speechivs) == 0:
            logging.warning("No speech segments in %s overlap with %s" % (filtertiername, speechseg))
            continue

        dbfilteredivs = tgt.IntervalTier()
        for siv in [x for x in speechivs if x.as_db > dbfilterthreshold]:
            dbfilteredivs.add_annotation(siv)
        stats_filtered += len(speechivs) - len(dbfilteredivs)
        stats_all += len(speechivs)
        if len(dbfilteredivs) == 0:
            logging.warning("All speech segments in %s dropped since their energy is below %.2f" % (
                speechseg, dbfilterthreshold))
            continue
        start_time = min([x.start_time for x in dbfilteredivs])
        end_time = max([x.end_time for x in dbfilteredivs])

        resulttier.add_annotation(tgt.Interval(
            start_time + shiftonset,
            end_time + shiftoffset, "speech"))
    assert stats_all > 0, "VAD was unable segment speech. Check silence region: calculated threshold %.2f db. " \
                          "Speech threshold: %.2f db." % (silencelevel, dbfilterthreshold)
    tier = resulttier
    logging.info("Dropped %d of %d speech segments (%.2f%%) with energy below %.2f db" %
                 (stats_filtered, stats_all, (stats_filtered / stats_all * 100), dbfilterthreshold))

    tier.name = "seg.speech"
    tg.add_tier(tier)
    logging.info("Writing %s" % outfile)
    tgt.io.write_to_file(textgrid=tg, filename=outfile, format="long")


def add_tier(infile, outfile, wavfile, mode, sourcetier, filtertier, desttier, text, pattern):
    if mode in ['trim', 'copy']:
        assert sourcetier is not None, "source tier required for mode %s" % mode
    logging.info("Adding tier using mode: " + mode)
    duration, _ = util.get_wav_duration(wavfile)
    tg, tier = util.init_textgrid(infile, duration, desttier)
    duration, _ = util.get_wav_duration(wavfile)
    tier.start_time = 0
    tier.end_time = duration
    if mode == "trim":
        annotier = tg.get_tier_by_name(sourcetier)
        if len(annotier) > 0:
            tier.add_annotation(tgt.Annotation(annotier[0].start_time, annotier[-1].end_time, text))
    elif mode == "all":
        tier.add_annotation(tgt.Annotation(0, duration, text))
    elif mode == "copy":
        annotier = tg.get_tier_by_name(sourcetier)
        for iv in annotier:
            if re.match(pattern, iv.text) is not None:
                tier.add_annotation(tgt.Annotation(iv.start_time, iv.end_time, text))
    elif mode == "trimright":
        overlaptier = tg.get_tier_by_name(sourcetier)
        trimtier = tg.get_tier_by_name(filtertier)
        for overlapseg in overlaptier.intervals:
            if re.match(pattern, overlapseg.text) is not None:
                overlaps = trimtier.get_annotations_between_timepoints(overlapseg.start_time, overlapseg.end_time,
                                                                                left_overlap=True, right_overlap=True)
                if len(overlaps) > 0:
                    tier.add_annotation(tgt.Annotation(overlapseg.start_time,
                                                       min(overlaps[-1].end_time+0.5, overlapseg.end_time), text))
    tg.add_tier(tier)
    logging.info("Writing %s" % outfile)
    tgt.io.write_to_file(textgrid=tg, filename=outfile, format="long")


def load_dict(filename):
    pdict = {}
    with open(filename, 'r') as dictfile:
        dictreader = csv.reader(dictfile, delimiter=';', quoting=csv.QUOTE_NONE)
        for row in dictreader:
            pdict[row[0]] = row[1]
    return pdict


def save_dict(pdict, filename):
    with open(filename, 'w') as dictfile:
        dictwriter = csv.writer(dictfile, delimiter=';', quoting=csv.QUOTE_NONE, lineterminator="\n")
        for key, val in pdict.items():
            dictwriter.writerow([key, val])


def get_phonetic_transcriptions(tmpdir, segtier, annotier, language):
    intervalcnt = 0
    logging.info("Preparing transcription dictionary")
    pdict = {}
    cachefilename = os.path.join(cachedir, "aligntool.%s.cache" % language)
    if os.path.exists(cachefilename):
        pdict = load_dict(cachefilename)

    missing_kan = set()
    for speechseg in segtier.intervals:
        if speechseg.text == "speech":
            intervalcnt += 1
            wordsegments = annotier.get_annotations_between_timepoints(speechseg.start_time, speechseg.end_time,
                                                                       left_overlap=True, right_overlap=True)
            if len(wordsegments) == 0:
                continue
            annotation = " ".join([x.text for x in wordsegments]).split()
            for w in annotation:
                assert "_" not in w, "Word %s contains invalid character _" % w
                if w not in pdict.keys():
                    missing_kan.add(w)
    logging.info("%s missing phonetic transcriptions" % len(missing_kan))
    if len(missing_kan) > 0:
        lexfile = os.path.join(tmpdir, "lexicon.txt")
        txtout = open(lexfile, 'w')
        for w in missing_kan:
            print(w, file=txtout)
        txtout.close()
        g2pscript = os.path.join(os.path.dirname(sys.argv[0]), "rung2pwebservice.sh")
        util.call_check([g2pscript, lexfile, language])
        lextransfile = os.path.join(tmpdir, "lexicon.tab")
        newdictlines = open(lextransfile, 'r').readlines()
        for w, response in zip(missing_kan, newdictlines):
            responsew, responset = response[:-1].split(';')
            assert responset != "", "word '%s' was mapped to empty string. Invalid chars?" % w
            if w != responsew:
                logging.warning("g2p expanded %s to %s" % (w, responsew))
            pdict[w] = "".join(responset.split(" "))
        save_dict(pdict, cachefilename)
    return pdict


def split_utterances(tmpdir, speechtiername, infile, wavfile, channel, denoise):
    logging.info("Splitting audio into utterance segments")
    splitaudioscript = os.path.join(os.path.dirname(sys.argv[0]), "splitaudio.praat")
    if denoise:
        logging.warning("Assuming 1-4 seconds are non-speech for denoising")
    result = util.call_check(["praat", "--run", splitaudioscript, os.path.realpath(wavfile), os.path.realpath(infile),
                              "%s/iv" % tmpdir, str(channel), speechtiername, str(int(denoise))], True)
    offsets = []
    for line in result.decode().split("\n"):
        items = line.split("\t")
        if len(items) == 3:
            foffset = tgt.Interval(start_time=float(items[0]), end_time=float(items[1]), text=items[2])
            offsets.append(foffset)
    logging.info("Split completed: %s segments" % len(offsets))
    return offsets


def parse_maus_par(parfilename, sample_rate):
    ort_ivs = []
    mau_ivs = []
    with open(parfilename, 'r') as parfile:
        # print(parfilename)
        parreader = csv.reader(parfile, delimiter='\t', quotechar=None)
        for row in parreader:
            if row[0] == "ORT:":
                oiv = tgt.Interval(0, 0, row[2])
                oiv.has_begin_set = False
                ort_ivs.append(oiv)
                assert len(ort_ivs) == int(row[1]) + 1
            elif row[0] == "MAU:":
                ivbegin = float(row[1]) / sample_rate
                ivend = (float(row[1]) + float(row[2]) + 1) / sample_rate
                wnum = int(row[3])
                # print(wnum, ivbegin, ivend, row[4])
                ort_ivs[wnum].end_time = ivend
                if wnum >= 0 and not ort_ivs[wnum].has_begin_set:
                    ort_ivs[wnum].start_time = ivbegin
                    ort_ivs[wnum].has_begin_set = True
                mau_ivs.append(tgt.Interval(ivbegin, ivend, row[4]))
    if not mau_ivs:
        return [], []
    for iv in ort_ivs:
        assert iv.has_begin_set, "Incomplete MAU tier in %s" % parfilename
    return ort_ivs, mau_ivs


def read_maus_alignments(tmpdir, offsets, orttier, mautier, sample_rate):
    logging.info("Reading MAUS alignments")
    for i, foffset in enumerate(offsets):
        intervalcnt = i+1
        parfile = "%s/iv%s.par" % (tmpdir, intervalcnt)
        try:
            ort_ivs, mau_ivs = parse_maus_par(parfile, sample_rate)
            if not ort_ivs and foffset.transcription_valid:
                logging.warning("No alignment imported for interval %s: %s" % (intervalcnt, foffset))
            for iv in ort_ivs:
                orttier.add_annotation(tgt.Interval(iv.start_time + foffset.start_time,
                                                    iv.end_time + foffset.start_time, iv.text))
            for iv in mau_ivs:
                mautier.add_annotation(tgt.Interval(iv.start_time + foffset.start_time,
                                                    iv.end_time + foffset.start_time, iv.text))
        except IOError:
            if foffset.transcription_valid:
                logging.warning("No alignment imported for interval %s: %s" % (intervalcnt, foffset))
        except:
            logging.error("Exception while parsing TextGrid %s" % parfile)
            raise


def generate_maus_transcriptions(tmpdir, segtier, offsets, annotier, pdict):
    logging.info("Generating transcriptions for MAUS alignment")
    for intervalcnt, foffset in enumerate(offsets):
        foffset.transcription_valid = False
        seg_ivs = segtier.get_annotations_between_timepoints(foffset.start_time, foffset.end_time,
                                                             left_overlap=True, right_overlap=True)
        seg_ivs = [x for x in seg_ivs if x.text == "speech"]
        assert len(seg_ivs) <= 1, "Invalid segmentation hierarchy: %s is overlap-contained by multiple " \
                                  "pre-segmentation intervals %s" % (foffset, seg_ivs)
        if not seg_ivs:
            logging.warning("%s does not seem to correspond to any pre-segmentation interval" % foffset)
            continue
        seg_iv = seg_ivs[0]
        assert seg_iv.text == "speech", "speech interval %s contained by non-speech pre-segmentation %s" % (foffset, seg_iv)
        wordsegments = annotier.get_annotations_between_timepoints(seg_iv.start_time, seg_iv.end_time,
                                                                   left_overlap=True, right_overlap=True)
        if len(wordsegments) == 0:
            logging.warning("Ignoring empty annotation for interval %s" % seg_iv)
            continue
        foffset.transcription_valid = True
        txtout = open("%s/iv%s.par" % (tmpdir, intervalcnt+1), 'w')
        words = " ".join([x.text for x in wordsegments]).split()
        for i, w in enumerate(words):
            print("ORT:\t%s\t%s" % (i, w), file=txtout)
        for i, w in enumerate(words):
            print("KAN:\t%s\t%s" % (i, pdict[w]), file=txtout)
        txtout.close()


def align_maus(infile, wavfile, outfile, denoise, channel, segtiername, filtertiername, initialsilence, remote,
               language):
    logging.info("Aligning %s based on segmentation in %s" % (wavfile, infile))
    tmpdir = tempfile.mkdtemp()
    try:
        duration, _ = util.get_wav_duration(wavfile)
        tg, orttier, mautier = util.init_textgrid(infile, duration, "maus.ort", "maus.pho")
        annotier = tg.get_tier_by_name("anno.trans")
        segtier = tg.get_tier_by_name(filtertiername)

        pdict = get_phonetic_transcriptions(tmpdir, segtier, annotier, language)
        offsets = split_utterances(tmpdir, segtiername, infile, wavfile, channel, denoise)
        generate_maus_transcriptions(tmpdir, segtier, offsets, annotier, pdict)

        logging.info("Performing MAUS alignment")
        if not os.path.exists(os.path.join(os.path.dirname(sys.argv[0]), "..", "external", "maus")):
            mausscript = os.path.join(os.path.dirname(sys.argv[0]), "runmauswebservice.sh")
        else:
            mausscript = os.path.join(os.path.dirname(sys.argv[0]), "runmauslocal.sh")
        util.call_check([mausscript, tmpdir, str(not initialsilence), language])

        duration, sample_rate = util.get_wav_duration(wavfile)
        read_maus_alignments(tmpdir, offsets, orttier, mautier, sample_rate)
        tg.add_tier(orttier)
        tg.add_tier(mautier)
        logging.info("Writing %s" % outfile)
        tgt.io.write_to_file(textgrid=tg, filename=outfile, format="long")
    except:
        logging.error("Exception while running maus alignment. Retained temp dir: %s" % tmpdir)
        raise
    else:
        shutil.rmtree(tmpdir)


class DumpRecord:
    pass


def dump_boundaries(infile, outfile, filtertiername, ref_seg_tiername):
    logging.info("Dumping boundaries in %s to %s" % (infile, outfile))
    tg = tgt.io.read_textgrid(infile)
    anno_tier = tg.get_tier_by_name("anno.trans")
    meta_tier = tg.get_tier_by_name("anno.meta")
    align_tier = tg.get_tier_by_name("maus.ort")
    pho_tier = tg.get_tier_by_name("maus.pho")
    seg_tier = tg.get_tier_by_name(filtertiername)
    if ref_seg_tiername is not None:
        ref_seg_tier = tg.get_tier_by_name(ref_seg_tiername)
    else:
        ref_seg_tier = None
    with open(outfile, "w") as fout:
        varnames = ("meta_ref", "meta_valid", "meta_orig", "meta_edited", "meta_ignore", "meta_shortened",
                    "meta_error", "transcription", "anno_onset", "anno_offset", "align_onset", "align_offset",
                    "word_cnt", "nwords", "onset_pho", "offset_pho", "anno_preseg_begin", "preseg_begin")
        print("\t".join(varnames), file=fout)
        # loop over utterances
        # loop over meta intervals
        # all valid intervals must have an matching anno.trans interval (ground truth)
        # hard index correspondence between maus.ort intervals and anno.trans intervals
        # one line output for each valid meta interval to catch edits and ignores
        for seg_iv in [s for s in seg_tier if s.text == "speech"]:
            word_cnt = 1
            anno_ivs = anno_tier.get_annotations_between_timepoints(seg_iv.start_time, seg_iv.end_time,
                                                                    left_overlap=True, right_overlap=True)
            meta_ivs = meta_tier.get_annotations_between_timepoints(seg_iv.start_time, seg_iv.end_time,
                                                                    left_overlap=True, right_overlap=True)
            align_ivs = align_tier.get_annotations_between_timepoints(seg_iv.start_time, seg_iv.end_time,
                                                                      left_overlap=True, right_overlap=True)
            ref_seg_iv = None
            if ref_seg_tier is not None:
                ref_seg_ivs = ref_seg_tier.get_annotations_between_timepoints(seg_iv.start_time, seg_iv.end_time,
                                                                              left_overlap=True, right_overlap=True)
                ref_seg_ivs = [s for s in ref_seg_ivs if s.text == "speech"]
                assert len(ref_seg_ivs) == 1, "unclear mapping to reference pre-segmentation %s" % ref_seg_ivs
                ref_seg_iv = ref_seg_ivs[0]

            anno_iter, align_iter = [iter(x) for x in (anno_ivs, align_ivs)]
            for metaiv in meta_ivs:
                record = DumpRecord()
                record.nwords = len(anno_ivs)
                metadata = util.MetaData.from_json(metaiv.text)
                record.meta_ref = metadata.ref
                record.meta_valid = int(metadata.valid)
                record.meta_orig = metadata.orig
                record.meta_edited = int(metadata.edited)
                record.meta_ignore = int(metadata.ignore)
                record.meta_shortened = int(metadata.shortened)
                record.meta_error = 0
                record.anno_onset = metaiv.start_time
                record.anno_offset = metaiv.end_time
                record.preseg_begin = seg_iv.start_time
                if ref_seg_iv is not None:
                    record.anno_preseg_begin = ref_seg_iv.start_time
                if metadata.valid:
                    try:
                        annoiv, aligniv = [next(x) for x in (anno_iter, align_iter)]
                        assert annoiv.start_time == metaiv.start_time and annoiv.end_time == metaiv.end_time, \
                            ("Metadata interval %s not matching annotation interval %s", metaiv, annoiv)
                        if annoiv.text != aligniv.text:
                            logging.warning(
                                "Mismatch between annotation %s and alignment result %s" % (annoiv, aligniv))
                        record.transcription = annoiv.text
                        record.align_onset = aligniv.start_time
                        record.align_offset = aligniv.end_time
                        pho_ivs = pho_tier.get_annotations_between_timepoints(aligniv.start_time, aligniv.end_time,
                                                                              left_overlap=True, right_overlap=True)
                        pho_ivs_cleaned = [x.text for x in pho_ivs if x.text not in ("?", "<p:>", "<usb>")]
                        if pho_ivs_cleaned:
                            record.onset_pho = pho_ivs_cleaned[0]
                            record.offset_pho = pho_ivs_cleaned[-1]
                        else:
                            logging.warning("Phonetic alignment for %s seems empty after filtering: %s"
                                            % (aligniv, pho_ivs_cleaned))
                            record.meta_error = 1
                        record.word_cnt = word_cnt
                        word_cnt += 1
                    except StopIteration:
                        logging.warning("Alignment in %s is missing %s" % (seg_iv, metadata.orig))
                        record.meta_error = 1
                print("\t".join([str(getattr(record, x, "")) for x in varnames]), file=fout)


def export_boundaries(infile, outfile, filtertiername):
    logging.info("Exporting boundaries in %s to %s" % (infile, outfile))
    tg = tgt.io.read_textgrid(infile)
    align_tier = tg.get_tier_by_name("maus.ort")
    pho_tier = tg.get_tier_by_name("maus.pho")
    seg_tier = tg.get_tier_by_name(filtertiername)
    with open(outfile, "w") as fout:
        varnames = ("meta_error", "transcription", "align_onset", "align_offset", "word_cnt", "nwords", "onset_pho",
                    "offset_pho", "preseg_begin")
        print("\t".join(varnames), file=fout)
        for seg_iv in [s for s in seg_tier if s.text == "speech"]:
            word_cnt = 1
            align_ivs = align_tier.get_annotations_between_timepoints(seg_iv.start_time, seg_iv.end_time,
                                                                      left_overlap=True, right_overlap=True)
            for aligniv in align_ivs:
                record = DumpRecord()
                record.transcription = aligniv.text
                record.nwords = len(align_ivs)
                record.meta_error = 0
                record.preseg_begin = seg_iv.start_time
                record.align_onset = aligniv.start_time
                record.align_offset = aligniv.end_time
                pho_ivs = pho_tier.get_annotations_between_timepoints(aligniv.start_time, aligniv.end_time,
                                                                      left_overlap=True, right_overlap=True)
                pho_ivs_cleaned = [x.text for x in pho_ivs if x.text not in ("?", "<p:>", "<usb>")]
                if pho_ivs_cleaned:
                    record.onset_pho = pho_ivs_cleaned[0]
                    record.offset_pho = pho_ivs_cleaned[-1]
                else:
                    logging.warning("Phonetic alignment for %s seems empty after filtering: %s"
                                    % (aligniv, pho_ivs_cleaned))
                    record.meta_error = 1
                record.word_cnt = word_cnt
                word_cnt += 1
                print("\t".join([str(getattr(record, x, "")) for x in varnames]), file=fout)


def add_textgrid_options(parser):
    parser.add_argument('-i', "--input-textgrid", dest='infile', metavar='<infile>', action='store', required=False,
                        help='input TextGrid')
    parser.add_argument('-o', "--output-textgrid", dest='outfile', metavar='<outfile>', action='store', required=True,
                        help='output TextGrid')


def add_language_options(parser):
    langs = ["deu-DE", "gsw-CH", "eng-GB", "fin-FI", "fra-FR", "eng-AU", "eng-US",
             "nld-NL", "spa-ES", "ita-IT", "por-PT", "hun-HU", "ekk-EE", "pol-PL",
             "eng-NZ", "kat-GE", "rus-RU"]
    parser.add_argument('-l', "--language", dest='language', metavar='<language>', action='store',
                        required=False, choices=langs, default="deu-DE", help='language code')


def parse_arguments(argv, parser=argparse.ArgumentParser(prog=os.path.basename(__file__), add_help=True)):
    sub_cmd_parser = parser.add_subparsers(dest='cmd', title='subcommands (-h for more help)')
    sub_cmd_parser.required = True

    xlsbatch_parser = sub_cmd_parser.add_parser('xlsbatch', help='batch process xls files ...',
                                                formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    sub_xlsbatch_parser = xlsbatch_parser.add_subparsers(dest='cmd', title='import commands')
    sub_xlsbatch_parser.required = True
    xlsbatch.setup(sub_xlsbatch_parser)

    segment_beep_parser = sub_cmd_parser.add_parser('segmentBeeps', help='segment audio files on beeps',
                                                    formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    segment_beep_parser.set_defaults(cmd=segment_beeps)
    add_textgrid_options(segment_beep_parser)
    segment_beep_parser.add_argument('-w', "--wav-file", dest='wavfile', metavar='<wavfile>', action='store',
                                     required=True,
                                     help='input wav file')
    segment_beep_parser.add_argument('-b', dest='beepchannel', type=int, default=2, metavar='<channel>', action='store',
                                     required=False, help='beep channel')
    segment_beep_parser.add_argument('-r', "--refbeep", dest='refbeep', metavar='<refbeep>', action='append',
                                     required=True,
                                     help='reference beep signal')
    segment_beep_parser.add_argument('-x', "--mincorrelation", dest='mincorrelation', type=float, default=0,
                                     metavar='<0-1>', action='store', help='minimum required cross correlation peak')
    segment_beep_parser.add_argument('-s', "--silencethreshold", dest='silencethreshold', type=float, default=-25,
                                     metavar='<db>', action='store',
                                     help='silence threshold for beep segmentation (default: -25)')
    segment_beep_parser.add_argument('-m', "--minsounding", dest='minsoundingduration', type=float, default=0.18,
                                     metavar='<s>', action='store', help='min sounding duration')
    segment_beep_parser.add_argument('--seekflank', dest='seekflank', action='store_true',
                                     help='enable seekflank heuristic for postprocessing the correlation method')

    segment_speech_parser = sub_cmd_parser.add_parser('segmentSpeech',
                                                      help='segment audio file into speech and non-speech',
                                                      formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    segment_speech_parser.set_defaults(cmd=segment_speech)
    add_textgrid_options(segment_speech_parser)
    segment_speech_parser.add_argument('-w', "--wav-file", dest='wavfile', metavar='<wavfile>', action='store',
                                       required=True,
                                       help='input wav file')
    segment_speech_parser.add_argument('-c', "--speech-channel", dest='channel', type=int, default=1,
                                       metavar='<channel>', action='store',
                                       required=False, help='speech channel')
    segment_speech_parser.add_argument('-f', "--filter-tier", dest='filtertiername', metavar='<tier>', action='store',
                                       required=False,
                                       help='filter/suppress speech intervals based on existing speech interval tier')
    segment_speech_parser.add_argument("-d", "--denoise", dest='denoise', action='store_true', help='enable denoising')
    segment_speech_parser.add_argument('--shiftonsets', dest='shiftonset', metavar='<s>', action='store', type=float,
                                       default=0,
                                       required=False, help='shift detected onsets by <s> seconds')
    segment_speech_parser.add_argument('--shiftoffsets', dest='shiftoffset', metavar='<s>', action='store', type=float,
                                       default=0,
                                       required=False, help='shift detected offsets by <s> seconds')

    segment_speech_parser.add_argument('--trainbegin', dest='trainbegin', metavar='<s>', action='store', type=float,
                                       default=3.3,
                                       required=False,
                                       help='interval for calculating reference energy ends at <s> seconds')
    segment_speech_parser.add_argument('--trainwindow', dest='trainwindow', metavar='<s>', action='store', type=float,
                                       default=1,
                                       required=False,
                                       help='interval for calculating reference energy ends at <s> seconds')
    segment_speech_parser.add_argument('--speechthresh', dest='speechthresh', metavar='<f>', action='store', type=float,
                                       default=0.5,
                                       required=False,
                                       help="Relation of the average interval intensity to the valid speech interval "
                                            "intensity")
    segment_speech_parser.add_argument('--snradd', dest='snradd', metavar='<db>', action='store', type=float,
                                       default=1,
                                       required=False, help="add <db> to the silence threshold")

    add_tier_parser = sub_cmd_parser.add_parser('addTier',
                                                help='add a tier, e.g. to mark the whole utterance as speech',
                                                formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    add_tier_parser.set_defaults(cmd=add_tier)
    add_tier_parser.add_argument('-m', "--mode", dest='mode', metavar='<mode>',
                                 action='store', required=True, choices=['all', 'trim', 'copy', 'trimright'],
                                 help='set mode <mode> (all, trim, copy)')
    add_tier_parser.add_argument('-w', "--wav-file", dest='wavfile', metavar='<wavfile>', action='store',
                                 required=True, help='input wav file')
    add_tier_parser.add_argument('-s', "--source-tier", dest='sourcetier', metavar='<tier>',
                                 action='store', default=None, required=False,
                                 help='use <tier> as source for copying or trimming')
    add_tier_parser.add_argument('-d', "--dest-tier", dest='desttier', metavar='<tier>',
                                 action='store', required=False, default="seg.speech",
                                 help='destination tier <tier>')
    add_tier_parser.add_argument('-t', "--text", dest='text', metavar='<text>',
                                 action='store', required=False, default="speech",
                                 help="set new tier's intervals content to <text>")
    add_tier_parser.add_argument('-f', "--filter", dest='pattern', metavar='<re>',
                                 action='store', required=False, default="speech",
                                 help="only copy intervals with content matching pattern <re>")
    add_tier_parser.add_argument("--filter-tier", dest='filtertier', metavar='<tier>',
                                 action='store', default=None, required=False,
                                 help='use <tier> as source for overlap based trimming')
    add_textgrid_options(add_tier_parser)

    align_maus_parser = sub_cmd_parser.add_parser('alignMAUS',
                                                  help='align files using the MAUS webservice',
                                                  formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    align_maus_parser.set_defaults(cmd=align_maus)
    add_textgrid_options(align_maus_parser)
    add_language_options(align_maus_parser)

    align_maus_parser.add_argument('-w', "--wav-file", dest='wavfile', metavar='<wavfile>', action='store',
                                   required=True,
                                   help='input wav file')
    align_maus_parser.add_argument('-c', "--speech-channel", dest='channel', metavar='<channel>', action='store',
                                   type=int, default=1,
                                   help='input wav file')
    align_maus_parser.add_argument('-f', "--filter-tier", dest='filtertiername', metavar='<tier>', action='store',
                                   default="seg.beep",
                                   help='<tier> providing initial segmentation, overlapping ground truth')
    align_maus_parser.add_argument('-s', "--segmentation-tier", dest='segtiername', metavar='<tier>', action='store',
                                   default="seg.beep",
                                   help='<tier> providing speech segmentation')
    align_maus_parser.add_argument("-d", "--denoise", dest='denoise', action='store_true', help='enable denoising')
    align_maus_parser.add_argument("--initialsilence", dest='initialsilence', action='store_true',
                                   help='enable initial and final silence models')
    align_maus_parser.add_argument("--remote", dest='remote', action='store_true',
                                   help='use maus online service')

    dump_boundaries_parser = sub_cmd_parser.add_parser('dumpBoundaries',
                                                       help='dump alignment and reference alignments for evaluation',
                                                       formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    dump_boundaries_parser.set_defaults(cmd=dump_boundaries)
    dump_boundaries_parser.add_argument('-i', "--input-textgrid", dest='infile', metavar='<infile>', action='store',
                                        required=False,
                                        help='input TextGrid')
    dump_boundaries_parser.add_argument('-o', "--output-file", dest='outfile', metavar='<outfile>', action='store',
                                        required=True,
                                        help='statistics file')
    dump_boundaries_parser.add_argument('-f', "--filter-tier", dest='filtertiername', metavar='<tier>', action='store',
                                        default="seg.beep",
                                        help='filter/suppress speech intervals based on existing speech interval tier')
    dump_boundaries_parser.add_argument('-r', "--ref_seg_tier", dest='ref_seg_tiername', metavar='<tier>', action='store',
                                        help='reference segmentation/filter tier')

    export_boundaries_parser = sub_cmd_parser.add_parser('exportBoundaries', help='export alignment results',
                                                         formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    export_boundaries_parser.set_defaults(cmd=export_boundaries)
    export_boundaries_parser.add_argument('-i', "--input-textgrid", dest='infile', metavar='<infile>', action='store',
                                          required=False,
                                          help='input TextGrid')
    export_boundaries_parser.add_argument('-o', "--output-file", dest='outfile', metavar='<outfile>', action='store',
                                          required=True,
                                          help='statistics file')
    export_boundaries_parser.add_argument('-f', "--filter-tier", dest='filtertiername', metavar='<tier>', action='store',
                                          default="seg.beep",
                                          help='pre-segmentation tier')
    gui_parser = sub_cmd_parser.add_parser('gui', help='open a simple graphical user interface',
                                           formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    gui_parser.set_defaults(cmd=gui.setup)

    args = parser.parse_args(argv)
    return args


def main():
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter('-=%(levelname)s=- [%(asctime)s.%(msecs)d] %(message)s', datefmt='%H:%M:%S')
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    try:
        args = parse_arguments(sys.argv[1:])
        args.cmd(**util.extract_args(args))
    except subprocess.CalledProcessError as e:
        print(sys.stderr, e)
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
