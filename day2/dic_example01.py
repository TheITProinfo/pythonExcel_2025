# utf-8
# Author: <NAME>
# Date: 2021-09-13
# this program demonstrates a dictionary in Python
person = {"name": "Alice", "age": 30, "city": "New York"}
print(person)  # Output: {'name': 'Alice', 'age': 30, 'city': 'New York'}
print(type(person))  # Output: <class 'dict'>
print(len(person))  # Output: 3
print(person["name"])  # Output: Alice
print(person["age"])  # Output: 30
print(person["city"])  # Output: New York
print(person["name"] + " is " + str(person["age"]) + " years old")  # Output: Alice is 30 years old
print(person["name"] + " lives in " + person["city"])  # Output: Alice lives in New York
print(person["name"] + " is " + str(person["age"]) + " years old and lives in " + person["city"])  # Output: Alice is 30 years old and lives in New York
print(person["name"] + " is " + str(person["age"]) + " years old and lives in " + person["city"])  # Output: Alice is 30 years old and lives in New York

for key in person:
    print(key, ":", person[key])    

    