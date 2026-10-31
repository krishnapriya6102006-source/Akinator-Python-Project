{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 1,
   "id": "019cce81-dda9-44ca-aec5-bbdddc928d84",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "==================================================\n",
      "             LIST OF PEOPLE\n",
      "==================================================\n",
      "1. Shah Rukh Khan\n",
      "2. A. R. Rahman\n",
      "3. Virat Kohli\n",
      "4. P. V. Sindhu\n",
      "5. C. V. Raman\n",
      "6. A. P. J. Abdul Kalam\n",
      "7. Deepika Padukone\n",
      "8. Neeraj Chopra\n",
      "9. Taylor Swift\n",
      "10. Beyoncé\n",
      "11. Tom Hanks\n",
      "12. Serena Williams\n",
      "13. Barack Obama\n",
      "14. Oprah Winfrey\n",
      "15. Cristiano Ronaldo\n",
      "16. Lionel Messi\n",
      "17. Angela Merkel\n",
      "18. Heidi Klum\n",
      "19. Marie Curie\n",
      "20. Adele\n",
      "21. David Attenborough\n",
      "22. Emma Watson\n",
      "23. Justin Bieber\n",
      "24. Jungkook\n",
      "25. Rafael Nadal\n",
      "26. Kylian Mbappé\n",
      "27. Neymar\n",
      "==================================================\n",
      "\n",
      "Welcome to Akinator!\n",
      "Think of a person from the list above.\n",
      "Answer using yes or no.\n",
      "\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person Indian? (yes/no):  yes\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 8\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person male? (yes/no):  yes\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 6\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person alive? (yes/no):  no\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 2\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person an actor? (yes/no):  no\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 2\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person a singer? (yes/no):  no\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 2\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person a sportsperson? (yes/no):  no\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 2\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person a scientist? (yes/no):  yes\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 2\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "Is the person a political leader? (yes/no):  no\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Possible people remaining: 1\n",
      "\n",
      "==================================================\n",
      "My guess is: C. V. Raman\n",
      "==================================================\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "\n",
      "Is my guess correct? (yes/no):  no\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Sorry! My guess was incorrect.\n",
      "I couldn't identify the correct person from the remaining data.\n",
      "\n",
      "Other people in the database you might mean:\n",
      "- Shah Rukh Khan\n",
      "- A. R. Rahman\n",
      "- Virat Kohli\n",
      "- P. V. Sindhu\n",
      "- A. P. J. Abdul Kalam\n",
      "- Deepika Padukone\n",
      "- Neeraj Chopra\n",
      "- Taylor Swift\n",
      "- Beyoncé\n",
      "- Tom Hanks\n",
      "- Serena Williams\n",
      "- Barack Obama\n",
      "- Oprah Winfrey\n",
      "- Cristiano Ronaldo\n",
      "- Lionel Messi\n",
      "- Angela Merkel\n",
      "- Heidi Klum\n",
      "- Marie Curie\n",
      "- Adele\n",
      "- David Attenborough\n",
      "- Emma Watson\n",
      "- Justin Bieber\n",
      "- Jungkook\n",
      "- Rafael Nadal\n",
      "- Kylian Mbappé\n",
      "- Neymar\n"
     ]
    },
    {
     "name": "stdin",
     "output_type": "stream",
     "text": [
      "\n",
      "Would you like to play again? (yes/no):  no\n"
     ]
    },
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Thanks for playing!\n"
     ]
    }
   ],
   "source": [
    "import pandas as pd\n",
    "\n",
    "# ============================================================\n",
    "# READ THE EXCEL FILE\n",
    "# ============================================================\n",
    "\n",
    "people = pd.read_excel(\"people_25.xlsx\")\n",
    "\n",
    "# Convert Yes/No values into a consistent format\n",
    "for column in people.columns[1:]:\n",
    "    people[column] = (\n",
    "        people[column]\n",
    "        .astype(str)\n",
    "        .str.strip()\n",
    "        .str.title()\n",
    "    )\n",
    "\n",
    "\n",
    "# ============================================================\n",
    "# DISPLAY THE LIST OF PEOPLE\n",
    "# ============================================================\n",
    "\n",
    "print(\"=\" * 50)\n",
    "print(\"             LIST OF PEOPLE\")\n",
    "print(\"=\" * 50)\n",
    "\n",
    "for i, person in enumerate(people[\"Person\"], start=1):\n",
    "    print(f\"{i}. {person}\")\n",
    "\n",
    "print(\"=\" * 50)\n",
    "\n",
    "\n",
    "# ============================================================\n",
    "# ASK QUESTION AND FILTER PEOPLE\n",
    "# ============================================================\n",
    "\n",
    "def ask_question(question, column, possible_people):\n",
    "    answer = input(question + \" (yes/no): \").strip().lower()\n",
    "\n",
    "    while answer not in [\"yes\", \"y\", \"no\", \"n\"]:\n",
    "        answer = input(\"Please enter yes or no: \").strip().lower()\n",
    "\n",
    "    expected_answer = \"Yes\" if answer in [\"yes\", \"y\"] else \"No\"\n",
    "\n",
    "    possible_people = possible_people[\n",
    "        possible_people[column] == expected_answer\n",
    "    ]\n",
    "\n",
    "    print(\"\\nPossible people remaining:\", len(possible_people))\n",
    "\n",
    "    return possible_people, expected_answer\n",
    "\n",
    "\n",
    "# ============================================================\n",
    "# AKINATOR GAME\n",
    "# ============================================================\n",
    "\n",
    "def play_game():\n",
    "\n",
    "    possible_people = people.copy()\n",
    "\n",
    "    print(\"\\nWelcome to Akinator!\")\n",
    "    print(\"Think of a person from the list above.\")\n",
    "    print(\"Answer using yes or no.\\n\")\n",
    "\n",
    "    # --------------------------------------------------------\n",
    "    # QUESTIONS\n",
    "    # --------------------------------------------------------\n",
    "\n",
    "    questions = [\n",
    "        (\"Is the person Indian?\", \"Indian\", \"nationality\"),\n",
    "        (\"Is the person American?\", \"American\", \"nationality\"),\n",
    "        (\"Is the person Portuguese?\", \"Portuguese\", \"nationality\"),\n",
    "        (\"Is the person Argentinian?\", \"Argentinian\", \"nationality\"),\n",
    "        (\"Is the person German?\", \"German\", \"nationality\"),\n",
    "        (\"Is the person Polish?\", \"Polish\", \"nationality\"),\n",
    "        (\"Is the person British?\", \"British\", \"nationality\"),\n",
    "        (\"Is the person Canadian?\", \"Canadian\", \"nationality\"),\n",
    "        (\"Is the person South Korean?\", \"South Korean\", \"nationality\"),\n",
    "        (\"Is the person Spanish?\", \"Spanish\", \"nationality\"),\n",
    "        (\"Is the person French?\", \"French\", \"nationality\"),\n",
    "        (\"Is the person Brazilian?\", \"Brazilian\", \"nationality\"),\n",
    "        (\"Is the person male?\", \"Male\", \"other\"),\n",
    "        (\"Is the person alive?\", \"Alive\", \"other\"),\n",
    "        (\"Is the person an actor?\", \"Actor\", \"other\"),\n",
    "        (\"Is the person a singer?\", \"Singer\", \"other\"),\n",
    "        (\"Is the person a sportsperson?\", \"Sportsperson\", \"other\"),\n",
    "        (\"Is the person a scientist?\", \"Scientist\", \"other\"),\n",
    "        (\"Is the person a political leader?\", \"Political Leader\", \"other\"),\n",
    "        (\"Is the person a footballer?\", \"Footballer\", \"other\"),\n",
    "        (\"Is the person known for movies?\", \"Known for Movies\", \"other\"),\n",
    "        (\"Was the person born before 1980?\", \"Born before 1980\", \"other\"),\n",
    "        (\"Has the person won an Olympic medal?\", \"Won Olympic Medal\", \"other\")\n",
    "    ]\n",
    "\n",
    "    asked_questions = set()\n",
    "    nationality_found = False\n",
    "\n",
    "    # --------------------------------------------------------\n",
    "    # ASK QUESTIONS\n",
    "    # --------------------------------------------------------\n",
    "\n",
    "    while len(possible_people) > 1:\n",
    "\n",
    "        available_questions = [\n",
    "            q for q in questions\n",
    "            if q[1] not in asked_questions\n",
    "            and (q[2] != \"nationality\" or not nationality_found)\n",
    "        ]\n",
    "\n",
    "        if not available_questions:\n",
    "            break\n",
    "\n",
    "        question, column, category = available_questions[0]\n",
    "\n",
    "        possible_people, answer = ask_question(\n",
    "            question, column, possible_people\n",
    "        )\n",
    "\n",
    "        asked_questions.add(column)\n",
    "\n",
    "        if category == \"nationality\" and answer == \"Yes\":\n",
    "            nationality_found = True\n",
    "\n",
    "        if len(possible_people) == 0:\n",
    "            print(\"\\nNo person matches your answers.\")\n",
    "            return\n",
    "\n",
    "    # --------------------------------------------------------\n",
    "    # MAKE A GUESS\n",
    "    # --------------------------------------------------------\n",
    "\n",
    "    if len(possible_people) == 1:\n",
    "\n",
    "        guessed_person = possible_people.iloc[0][\"Person\"]\n",
    "\n",
    "        print(\"\\n\" + \"=\" * 50)\n",
    "        print(\"My guess is:\", guessed_person)\n",
    "        print(\"=\" * 50)\n",
    "\n",
    "        while True:\n",
    "            correct = input(\n",
    "                \"\\nIs my guess correct? (yes/no): \"\n",
    "            ).strip().lower()\n",
    "\n",
    "            if correct in [\"yes\", \"y\", \"no\", \"n\"]:\n",
    "                break\n",
    "\n",
    "            print(\"Please enter yes or no.\")\n",
    "\n",
    "        if correct in [\"yes\", \"y\"]:\n",
    "            print(\"\\nGreat! I guessed your person correctly!\")\n",
    "\n",
    "        else:\n",
    "            print(\"\\nSorry! My guess was incorrect.\")\n",
    "            print(\"I couldn't identify the correct person from the remaining data.\")\n",
    "\n",
    "            print(\"\\nOther people in the database you might mean:\")\n",
    "\n",
    "            other_people = people[\n",
    "                people[\"Person\"] != guessed_person\n",
    "            ]\n",
    "\n",
    "            for person in other_people[\"Person\"]:\n",
    "                print(\"-\", person)\n",
    "\n",
    "    elif len(possible_people) > 1:\n",
    "\n",
    "        print(\"\\nI couldn't identify one exact person.\")\n",
    "        print(\"The possible people are:\")\n",
    "\n",
    "        for person in possible_people[\"Person\"]:\n",
    "            print(\"-\", person)\n",
    "\n",
    "        print(\"\\nTry adding more questions to identify them.\")\n",
    "\n",
    "    else:\n",
    "        print(\"\\nNo matching person found.\")\n",
    "\n",
    "\n",
    "# ============================================================\n",
    "# MAIN LOOP\n",
    "# ============================================================\n",
    "\n",
    "while True:\n",
    "\n",
    "    play_game()\n",
    "\n",
    "    again = input(\n",
    "        \"\\nWould you like to play again? (yes/no): \"\n",
    "    ).strip().lower()\n",
    "\n",
    "    if again not in [\"yes\", \"y\"]:\n",
    "        print(\"\\nThanks for playing!\")\n",
    "        break"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "5cd5b7ea-d167-4b7a-ad18-b69ccaf2b07f",
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.13.9"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
