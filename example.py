from GdocxHandler import *
import main

EXAMPLE_INPUT_FILE = "example.txt"
EXAMPLE_OUTPUT_FILE = "example.docx"
EXAMPLE_INPUT_STRING = f'''
({EchoHandler.NAME } "This line is printed during processing of {EXAMPLE_INPUT_FILE}. Эта строка выведена во время обработки {EXAMPLE_INPUT_FILE}")
({UnorderedListItemHandler.NAME}
    FIRST ITEM
)
({UnorderedListItemHandler.NAME}
    SECOND ITEM
)
({ParStyleHandler.NAME} heading-1
    BIG HEADER
)
({ParStyleHandler.NAME} heading-2
    Small header
)
({ParStyleHandler.NAME} heading-2
    Small header
)
# I can also process comments
# These lines won't be in final .docx document
({OrderedListHandler.NAME}
    ({OrderedListItemHandler.NAME}
        FIRST ITEM
    )
    ({OrderedListItemHandler.NAME}
        SECOND ITEM
    )
)
# Also I can put images and image captions
# Just uncomment these lines and replace IMAGE_PATH with path to an image:
#({ImageHandler.NAME} IMAGE_PATH)
#({ImageCaptionHandler.NAME}
#    This is image caption
#)

And we can also support regular text.
Isn't it wonderful?
'''

EXAMPLE_COMMENT = f'''Created files:
    {EXAMPLE_INPUT_FILE}
    {EXAMPLE_OUTPUT_FILE}
>> You can look into example.txt to see, how to write your own input files.
>> example.docx is the output file which was constructed from example.txt by the tool.

To use YOUR_INPUT_FILE_PATH.txt, use the following command:
    python3 main.py -i YOUR_INPUT_FILE_PATH.txt -o YOUR_OUTPUT_FILE_PATH.txt -s -se

Созданные файлы:
    {EXAMPLE_INPUT_FILE}
    {EXAMPLE_OUTPUT_FILE}
>> Ты можешь посмотреть содержимое example.txt, чтобы увидеть, как написать свой .txt файл.
>> example.docx это файл, который был создан из example.txt этим скриптом.

Чтобы использовать СВОЙ_ФАЙЛ.txt, введи следующую команду:
    python3 main.py -i СВОЙ_ФАЙЛ.txt -o СВОЙ_РЕЗ_ФАЙЛ.docx -s -se
'''

if __name__ == "__main__":
    print(EXAMPLE_COMMENT)

    inpath = EXAMPLE_INPUT_FILE
    outpath = EXAMPLE_OUTPUT_FILE
    file = open(inpath, "w")
    file.write(EXAMPLE_INPUT_STRING)
    file.close()

    GdocxParsing.INDENT_STRING = "    "
    GdocxParsing.STRIP_INDENT = True
    GdocxParsing.SKIP_EMPTY = True

    input("Press Enter to create .docx file\n")
    main.process_file(inpath, outpath)
else:
    print(f"{__name__} can't be used as module")
    exit(1)
