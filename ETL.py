{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 1,
   "id": "9a29b412",
   "metadata": {},
   "outputs": [],
   "source": [
    "import re\n",
    "import pandas as pd"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 3,
   "id": "8e20b588",
   "metadata": {},
   "outputs": [
    {
     "ename": "ValueError",
     "evalue": "Length mismatch: Expected axis has 18 elements, new values have 13 elements",
     "output_type": "error",
     "traceback": [
      "\u001b[0;31m---------------------------------------------------------------------------\u001b[0m",
      "\u001b[0;31mValueError\u001b[0m                                Traceback (most recent call last)",
      "\u001b[0;32m/var/folders/7q/ybsshgf94s772d_zxssm1wgw0000gn/T/ipykernel_77717/1971775611.py\u001b[0m in \u001b[0;36m<module>\u001b[0;34m\u001b[0m\n\u001b[1;32m     23\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     24\u001b[0m \u001b[0;31m# Set the DataFrame columns\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 25\u001b[0;31m \u001b[0mdf\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mcolumns\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mcolumns\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     26\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     27\u001b[0m \u001b[0;31m# Display the resulting DataFrame\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/opt/anaconda3/lib/python3.9/site-packages/pandas/core/generic.py\u001b[0m in \u001b[0;36m__setattr__\u001b[0;34m(self, name, value)\u001b[0m\n\u001b[1;32m   5586\u001b[0m         \u001b[0;32mtry\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m   5587\u001b[0m             \u001b[0mobject\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m__getattribute__\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mname\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m-> 5588\u001b[0;31m             \u001b[0;32mreturn\u001b[0m \u001b[0mobject\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m__setattr__\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mname\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mvalue\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m   5589\u001b[0m         \u001b[0;32mexcept\u001b[0m \u001b[0mAttributeError\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m   5590\u001b[0m             \u001b[0;32mpass\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/opt/anaconda3/lib/python3.9/site-packages/pandas/_libs/properties.pyx\u001b[0m in \u001b[0;36mpandas._libs.properties.AxisProperty.__set__\u001b[0;34m()\u001b[0m\n",
      "\u001b[0;32m~/opt/anaconda3/lib/python3.9/site-packages/pandas/core/generic.py\u001b[0m in \u001b[0;36m_set_axis\u001b[0;34m(self, axis, labels)\u001b[0m\n\u001b[1;32m    767\u001b[0m     \u001b[0;32mdef\u001b[0m \u001b[0m_set_axis\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0maxis\u001b[0m\u001b[0;34m:\u001b[0m \u001b[0mint\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mlabels\u001b[0m\u001b[0;34m:\u001b[0m \u001b[0mIndex\u001b[0m\u001b[0;34m)\u001b[0m \u001b[0;34m->\u001b[0m \u001b[0;32mNone\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    768\u001b[0m         \u001b[0mlabels\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mensure_index\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mlabels\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 769\u001b[0;31m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_mgr\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mset_axis\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0maxis\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mlabels\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    770\u001b[0m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_clear_item_cache\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    771\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/opt/anaconda3/lib/python3.9/site-packages/pandas/core/internals/managers.py\u001b[0m in \u001b[0;36mset_axis\u001b[0;34m(self, axis, new_labels)\u001b[0m\n\u001b[1;32m    212\u001b[0m     \u001b[0;32mdef\u001b[0m \u001b[0mset_axis\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0maxis\u001b[0m\u001b[0;34m:\u001b[0m \u001b[0mint\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mnew_labels\u001b[0m\u001b[0;34m:\u001b[0m \u001b[0mIndex\u001b[0m\u001b[0;34m)\u001b[0m \u001b[0;34m->\u001b[0m \u001b[0;32mNone\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    213\u001b[0m         \u001b[0;31m# Caller is responsible for ensuring we have an Index object.\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 214\u001b[0;31m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_validate_set_axis\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0maxis\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mnew_labels\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    215\u001b[0m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0maxes\u001b[0m\u001b[0;34m[\u001b[0m\u001b[0maxis\u001b[0m\u001b[0;34m]\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mnew_labels\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    216\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/opt/anaconda3/lib/python3.9/site-packages/pandas/core/internals/base.py\u001b[0m in \u001b[0;36m_validate_set_axis\u001b[0;34m(self, axis, new_labels)\u001b[0m\n\u001b[1;32m     67\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     68\u001b[0m         \u001b[0;32melif\u001b[0m \u001b[0mnew_len\u001b[0m \u001b[0;34m!=\u001b[0m \u001b[0mold_len\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 69\u001b[0;31m             raise ValueError(\n\u001b[0m\u001b[1;32m     70\u001b[0m                 \u001b[0;34mf\"Length mismatch: Expected axis has {old_len} elements, new \"\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     71\u001b[0m                 \u001b[0;34mf\"values have {new_len} elements\"\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;31mValueError\u001b[0m: Length mismatch: Expected axis has 18 elements, new values have 13 elements"
     ]
    }
   ],
   "source": [
    "\n",
    "# Read the Excel file\n",
    "df = pd.read_excel(\"1.xlsx\", header=None)\n",
    "\n",
    "# Transpose the DataFrame to convert rows into columns\n",
    "df = df.T\n",
    "\n",
    "# Define the column names\n",
    "columns = [\n",
    "    \"Research Group\",\n",
    "    \"Cross Reference\",\n",
    "    \"Research Item Topic\",\n",
    "    \"Source\",\n",
    "    \"Summary\",\n",
    "    \"Keywords\",\n",
    "    \"Relevant To\",\n",
    "    \"Primary Contact(s)\",\n",
    "    \"Others with Expertise\",\n",
    "    \"URL Links\",\n",
    "    \"Person Submitting Material\",\n",
    "    \"Submitted On\",\n",
    "    \"Supporting Documents\",\n",
    "]\n",
    "\n",
    "# Set the DataFrame columns\n",
    "df.columns = columns\n",
    "\n",
    "# Display the resulting DataFrame\n",
    "print(df)\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "ead96245",
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
   "version": "3.9.13"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
