# Place, Create and Remove Door Opening From A Linked Models
Extract location and rough opening size of doors from linked Revit model.  The data is filters by walls which only contain a structural layer. Translates the data from the linked model and creates doors openings in your current model.  

Using the point location of a door a method will find the nearest wall curve line to host the door.  The door is serialized using the linked model's element id in a parameter in your current model.  If the door moves it will compare element the element id and find it's new location.  If the element id isn't found in your current model it will generate a new door. If the element Id isn't found in the linked model the door will be deleted.  Door only used in your model can be removed from the list to compare I'm using the word "STRUC" to remove those doors from the list before beginning my comparisons.  If an opening size isn't available it will take a size 10 x 10 and duplicate it creating a new size of your door and making a new type labeling it using feet.

Keywords: Revit

## Installation
Copy .dll file and .addin to the [Revit Add-Ins folder](http://help.autodesk.com/view/RVT/2015/ENU/?guid=GUID-4FFDB03E-6936-417C-9772-8FC258A261F7).
Also look at my [DBLibrary](https://github.com/dannysbentley/DBLibrary) for missing methods. 

## Usage

Works in conjuntion with Revit. Developed with Revit API 2016.


## License

This sample is licensed under the terms of the [MIT License](https://opensource.org/licenses/MIT).

Copyright (c) 2017 Danny Bentley

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
