import aaf2
import pandas as pd

data = pd.read_excel(r"Provys_export.xlsx")
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
    "ComplianceTask": "PG",
    "EditorialTask": "EDT",
    "TransmissionCheckTask": "TRM",
}

edit_rate = 25
timecode_fps = 25

with aaf2.open("import_me.aaf", "w") as f:

    for index, row in df.iterrows():
        print("Processing ", row["Programme"])

        num_of_seg = row["Count of segments"]

        for i in range(num_of_seg):
            comp_mob = f.create.CompositionMob()
            comp_mob.usage = "Usage_TopLevel"
            comp_mob.name = row["Programme"] + " [" + str(i + 1).rjust(2, "0") + "]"

            tag = f.create.TaggedValue("ProvysTask", provystask_dict[row["Type"]])
            comp_mob["UserComments"].append(tag)

            tag = f.create.TaggedValue("ProvysVersion", row["Version"])
            comp_mob["UserComments"].append(tag)

            # tag = f.create.TaggedValue("Start", "10:00:00:00")
            # comp_mob["UserComments"].append(tag)

            tag = f.create.TaggedValue(
                "ProvysAudio", str(row["Completed audio"]).replace("nan", "")
            )
            comp_mob["UserComments"].append(tag)

            tag = f.create.TaggedValue(
                "ProvysSubs", str(row["Completed subs"]).replace("nan", "")
            )
            comp_mob["UserComments"].append(tag)

            tag = f.create.TaggedValue(
                "TapeID", row["House ID"].replace("_01", "_" + str(i + 1).rjust(2, "0")).replace("-01", "-" + str(i + 1).rjust(2, "0"))
            )
            comp_mob["UserComments"].append(tag)

            f.content.mobs.append(comp_mob)

            tc_slot = comp_mob.create_empty_sequence_slot(
                edit_rate, media_kind="timecode"
            )
            tc = f.create.Timecode(timecode_fps, True)
            tc.start = 900000
            tc_slot.segment.components.append(tc)

            nested_slotV = comp_mob.create_timeline_slot(edit_rate)
            nested_slotV["PhysicalTrackNumber"].value = 1
            nested_scopeV = f.create.NestedScope()
            nested_slotV.segment = nested_scopeV

            sequenceV = f.create.Sequence(media_kind="picture")
            nested_scopeV.slots.append(sequenceV)
            comp_fillV = f.create.Filler("picture", 0)
            sequenceV.components.append(comp_fillV)

            nested_slotA = comp_mob.create_timeline_slot(edit_rate)
            nested_slotA["PhysicalTrackNumber"].value = 1
            nested_scopeA = f.create.NestedScope()
            nested_slotA.segment = nested_scopeA

            for j in range(2):
                sequenceA = f.create.Sequence(media_kind="sound")
                nested_scopeA.slots.append(sequenceA)
                comp_fillA = f.create.Filler("sound", 0)
                sequenceA.components.append(comp_fillA)
