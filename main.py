import os
import re
import docx
import argparse


def initialization_command_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", type=str, required=True, help="dir path")
    parser.add_argument("--sp", type=str, required=True, help="list speakers")
    return parser.parse_args()


def process_speakers(arg_sp) -> list:
    if arg_sp.count(',') != 0:
        return arg_sp.split(',')
    else:
        return [arg_sp]


def definition_speaker(count, list_s) -> str:
    if len(list_s) == 1:
        return list_s[0]
    else:
        return list_s[count]


def count_video_one_w(text) -> int:
    types_video = re.findall(r"\s[dq]", text)
    return len(types_video)


def clearing_word(text) -> str:
    word_without_videoname = re.sub(r"\(.+\s*", "", text)
    word_without_spacing = re.findall(r"\s*(.+\b)", word_without_videoname)
    return word_without_spacing[0]


def create_list_words(lines, speaker) -> list:
    result = list()
    for line in lines:
        if len(line.text.replace(' ', '').replace('\n', '')) == 0:
            continue
        number_videos_for_one_word = count_video_one_w(line.text)
        clear_word = clearing_word(line.text)

        result.append(f"{clear_word.lower()}({number_videos_for_one_word} - {speaker})")
    return result


def print_word_list(words):
    print(", ".join(words))
    print('\n')


def processing_docx(name_dir, file, docx_speaker):
    with open(os.path.join(name_dir, file), 'rb') as f:
        document = docx.Document(f)
        list_words = create_list_words(document.paragraphs, docx_speaker)
        print_word_list(list_words)


def processing_dir(name_dir, sp):
    sorted_sir = sorted(os.listdir(name_dir))
    for count, file in enumerate(sorted_sir):
        if file.endswith(".docx"):
            speaker = definition_speaker(count-1, sp)
            processing_docx(name_dir, file, speaker)


def main():
    args = initialization_command_arguments()
    NAME_DIR = args.dir
    SPEAKERS = process_speakers(args.sp)
    # NAME_DIR = "files_words"
    # SPEAKERS = process_speakers("adam")

    if os.path.isdir(NAME_DIR):
        processing_dir(NAME_DIR, SPEAKERS)
    else:
        raise FileNotFoundError


if __name__ == "__main__":
    main()
