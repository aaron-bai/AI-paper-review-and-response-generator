# Automated Peer Review

## Dependencies

Install the required packages before running the script:

```bash
pip install -r requirements.txt
```

## Usage

Run the pipeline from the terminal as before. After the review and response text are generated, the program will prompt you to enter an output folder path in the command line, then save two Word documents in that folder:

- 评审意见+时间戳.docx
- 回复+时间戳.docx

If `--output-path` is provided, the combined plain-text output is still written to that file as well.