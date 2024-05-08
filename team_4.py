import pandas as pd
from colorama import Fore, Back, Style, init
import xlsxwriter
import os
import glob
import json
from datetime import datetime, timedelta

# Initialize Colorama
init(autoreset=True)

date_data = 'date_data.json'

def find_file(scores_files, user_input):
    for filename in scores_files:
        if filename.strip().lower().startswith(user_input):
            return filename
    return None


# Load the Excel file
filename = 'T4_Data.xlsx'
#df = pd.read_excel(filename)
df = pd.read_excel(filename)
scores_files = glob.glob('*_scores.xlsx')


if scores_files:
    files_list = '\n'.join(f"{index}: {file}" for index, file in enumerate(scores_files, start=1))
    #print("Found the following '_scores.xlsx' files:\n" + files_list)
    user_ans = input(Fore.CYAN + "Do you want to continue annotating an existing file? Y/N: ").strip().lower()
    if user_ans == 'y':
        print("\nFound the following '_scores.xlsx' files:\n" + files_list + '\n')
        for index, file in enumerate(scores_files, start=1):
            try:
                selected_index = int(input("Enter the number of the file you want to select: \n")) - 1
                # Validate the selected index
                if selected_index >= 0 and selected_index < len(scores_files):
                    selected_file = scores_files[selected_index]
                    #print(f"You have selected: {selected_file}.")
                    user_response = input(Fore.CYAN + "Continue annotating this file? (Y/N): " + selected_file + '\n').strip().lower()
                    if user_response == 'y':
                        existing_df = pd.read_excel(selected_file)
                        last_sentence = existing_df.iloc[-1]['Sentence']
                        print(last_sentence)
                        # Find this sentence in the shuffled DataFrame
                        start_index = df[df.iloc[:, 1].eq(last_sentence)].index[0] + 1
                        print(start_index)
                        break
                    else:
                        existing_df = pd.DataFrame(columns=['Sentence', 'Score'])
                        start_index = 0
                        initials = input(Fore.CYAN + "Enter your initials to create a new annotation file: ")
                        selected_file = initials + '_T4_scores.xlsx'
                        break
                    # Proceed with operations on the selected_file
                else:
                    print(Fore.RED +"Invalid selection. Please enter a valid number.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")
    else:
        existing_df = pd.DataFrame(columns=['Sentence', 'Score'])
        start_index = 0
        initials = input(Fore.CYAN + "Enter your initials: ")
        selected_file = initials + '_T4_scores.xlsx'
else:
    # No existing file, start fresh
    existing_df = pd.DataFrame(columns=['Sentence', 'Score'])
    start_index = 0
    selected_file = "new_scores.xlsx"




output_filename = selected_file
if output_filename == "new_scores.xlsx":
    # Ask for annotator's initials
    initials = input(Fore.CYAN + "Please enter your initials: ")
    output_filename = f"{initials}_T4_scores.xlsx"
    if output_filename in scores_files:
        output_filename = "new_" + output_filename

date_data_filename = f"{output_filename}_date_data.json"

WEEKLY_GOAL = 200

def ensure_file_exists(filename):
    """Check if the specific date_data file exists, create if not."""
    if not os.path.exists(filename):
        data = {'start_date': datetime.now().strftime('%Y-%m-%d'), 'annotations_completed': 0}
        with open(filename, 'w') as file:
            json.dump(data, file)

def read_data(filename):
    """Read data from the specific date_data file."""
    ensure_file_exists(filename)  # Make sure file exists
    with open(filename, 'r') as file:
        return json.load(file)

def write_data(filename, data):
    """Write data to the specific date_data file."""
    with open(filename, 'w') as file:
        json.dump(data, file)

def update_annotation_count(filename, change):
    """Update annotation count in the specific date_data file."""
    data = read_data(filename)
    data['annotations_completed'] += change
    write_data(filename, data)
    display_weekly_progress(filename)

def increment_annotation():
    """Increment annotation count for the specific annotator."""
    update_annotation_count(date_data_filename, 1)

def decrement_annotation():
    """Decrement annotation count for the specific annotator."""
    update_annotation_count(date_data_filename, -1)

def display_weekly_progress(filename):
    """Display remaining annotations for the week for the specific annotator."""
    data = read_data(filename)
    annotations_remaining = WEEKLY_GOAL - data['annotations_completed']
    if annotations_remaining >= 0:
        print(f"\n{Fore.RED}Annotations remaining this week: {annotations_remaining}/{WEEKLY_GOAL}")
    else:
        annotations_ahead = abs(annotations_remaining)
        print(f"\n{Fore.RED}Good job! You reached the weekly goal. You are {annotations_ahead} annotations ahead of the weekly target.")

# Function to ask for annotation, validate input, and allow re-answering previous questions
def annotate_sentences(df, resuming_index=0):
    results_df = pd.DataFrame(columns=['Sentence', 'Score'])
    index = resuming_index
    display_weekly_progress(date_data_filename)
    while index < len(df):
        row = df.iloc[index]
        sentence1, target_sentence, sentence3 = row[0], row[1], row[2]

        border_line = Fore.GREEN + "-" * 50  # Green line to simulate a border
        # Print the sentences with requested formatting and simulated green border
        print(border_line)
        print(sentence1 if pd.notna(sentence1) else "")
        print(f"{Fore.GREEN}{target_sentence}{Style.RESET_ALL}")  # Target sentence in green
        print(sentence3 if pd.notna(sentence3) else "")
        print(border_line)
        #print(f"Sentences remaining: {len(df) - index - 1}")

        question = Fore.CYAN + "Is the sentence in green anti-autistic? Y, N, C if it needs more context, or B to go back and re-label.\n"
        valid_responses = {'y': 1, 'yes': 1, 'n': 0, 'no': 0, 'c': -1, 'context': -1}
        while True:
            response = input(question + Style.RESET_ALL).strip().lower()  # Reset style after question
            if response in valid_responses:
                score = valid_responses[response]
                break
            elif response == 'b':
                decrement_annotation()
                if index == 0:
                    print(Fore.RED + "Error: You are at the first sentence. Cannot go back.")
                else:
                    index -= 1  # Move back to re-answer the previous question
                    if len(results_df) == 0:
                        existing_df.drop(existing_df.tail(1).index, inplace=True)
                    else:
                        results_df.drop(results_df.tail(1).index, inplace=True)
                    break
            elif response == 'q':
                print(Fore.CYAN + "We are ending the annotation.")
                return results_df
            else:
                print(Fore.RED + "Invalid character entered. Only Y, N, C, and B are acceptable answers.")

        if response != 'b' and response != 'q':  # Only add or update if not going back
            if index < len(results_df):
                results_df.at[index, 'Score'] = score
                final_df = results_df
            else:
                new_row = pd.DataFrame({'Sentence': [target_sentence], 'Score': [score]})
                results_df = pd.concat([results_df, new_row], ignore_index=True)
                final_df = pd.concat([existing_df, results_df], ignore_index=True, axis=0)
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False)
            index += 1
            increment_annotation()
        elif response == 'b':
            if index < len(results_df):
                final_df = results_df
            else:
                final_df = pd.concat([existing_df, results_df], ignore_index=True, axis=0)
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False)

    return results_df


# Start the annotation process
results_df = annotate_sentences(df, resuming_index=start_index)


print(Fore.RED + "Annotation complete. Please check to ensure your answers were saved properly. Results saved to: ", output_filename)