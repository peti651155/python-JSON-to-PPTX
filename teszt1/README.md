# Report generation application

## Description

This application is used to generate a report in pptx format based on a configuration file.

Module to be used:

- [pptx](https://python-pptx.readthedocs.io/en/latest/)

Please read the description of the assessment before starting the task.

## Task description

1. Create a command line application that accepts a configuration file as an argument and generates a report in pptx format.
2. The configuration file should be in json format *(see "sample.json"*).
3. 5 different types of slides should be supported *(hint: use slide layouts)*:
    - Title slide
    - Text slide
    - List slide
    - Picture slide
    - Plot slide (the plot data should be read from a .dat file)
4. The application should be able to generate a report with any number of slides in any order. An example of the generated report is shown in the file "example_output.pptx".

## Requirements

- Use Python 3.7 or higher
- Plot can be generated using any library (including the `Chart` module from `pptx`)
- The application should be able to run on Windows and Linux
- Use functions and classes where appropriate

### Optional requirements

- Use docstrings and comments where appropriate (Optional)
- Write unit tests (Optional)
- Use `numpy` for chart data (Optional)
- Have a logger and log important events (Optional)
- Handle possible exceptions (Optional)

## Additional information

- **Feel free to use any information available on the internet.**
- It's not necessary to implement all the requirements to submit the task, but try to show your best.
- If you have any questions, please contact us.
