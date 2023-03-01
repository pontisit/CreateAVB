import avb
import avb.utils
import pandas as pd
import sys
from datetime import datetime

ABC = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_-': /"
HUN = "áéíóöőõúüűûÁÉÍÓÖŐÕÚÜŰÛ"
ENG = "aeioooouuuuAEIOOOOUUUU"


def to_alpha_numeric(input_str):
    output_str = ""
    for char in input_str:
        hun_idx = HUN.find(char)
        if hun_idx >= 0:
            output_str += ENG[hun_idx]
        elif char not in ABC:
            output_str += "_"
        else:
            output_str += char
    return output_str


if len(sys.argv) < 2:
    print("Drag&Drop input excel!")
    pass
else:
    print("Processing input file ", sys.argv[1])

    data = pd.read_excel(sys.argv[1])
    df = pd.DataFrame(
        data,
        columns=[
            "Type",
            "Programme",
            "House ID",
            "Count of segments",
            "Version",
            "Completed audio",
            "Completed subs",
            "Details",
            "Directory",
            "Status",
        ],
    )

    provystask_dict = {
        "ComplianceTask": "MPG",
        "EditorialTask": "EDT",
        "TransmissionCheckTask": "TRM",
    }

    edit_rate = 25
    timecode_fps = 25
    dur = 25

    now = datetime.now()
    dt_string = now.strftime("_%Y-%m-%d_%H-%M-%S")

    result_file = sys.argv[1].replace(".xlsx", dt_string + ".avb")
    with avb.open() as f:

        for index, row in df.iterrows():
            print("Processing ", row["Programme"])

            num_of_seg = row["Count of segments"]

            for i in range(num_of_seg):

                comp = f.create.Composition(mob_type="CompositionMob")
                comp.name = to_alpha_numeric(row["Programme"]).upper()

                new_mob_att = f.create.Attributes()

                new_mob_att["ProvysTask"] = provystask_dict[row["Type"]]
                new_mob_att["ProvysVersion"] = row["Version"]
                new_mob_att["ProvysAudio"] = str(row["Completed audio"]).replace(
                    "nan", ""
                )
                new_mob_att["ProvysSubs"] = str(row["Completed subs"]).replace(
                    "nan", ""
                )
                new_mob_att["TapeID"] = row["House ID"].replace(
                    "_01", "_" + str(i + 1).rjust(2, "0")
                )

                comp.attributes["_USER"] = new_mob_att

                # timecode track
                track = f.create.Track()
                track.index = 1
                track.component = f.create.Timecode(
                    edit_rate=edit_rate, media_kind="timecode"
                )
                track.component.start = 900000
                track.component.fps = 25
                track.component.length = 0
                comp.tracks.append(track)

                # A1
                track = f.create.Track()
                track.index = 1
                sequence = f.create.Sequence(edit_rate=edit_rate, media_kind="sound")
                fill = f.create.Filler(edit_rate=edit_rate, media_kind="sound")
                fill.length = comp.length
                sequence.components.append(fill)
                track.component = sequence
                comp.tracks.append(track)

                # A2
                track = f.create.Track()
                track.index = 2
                sequence = f.create.Sequence(edit_rate=edit_rate, media_kind="sound")
                fill = f.create.Filler(edit_rate=edit_rate, media_kind="sound")
                fill.length = comp.length
                sequence.components.append(fill)
                track.component = sequence
                comp.tracks.append(track)

                f.content.add_mob(comp)

                f.write(result_file)
