# excel_vba_Perl_style_arrays


REQUIRES MSExcel 2016 or higher

Trying to do anything complex with VBA arrays is like banging your head on a wall but without that satisfying feeling when you stop. Because of this, I created this class to do the nifty stuff I can do MUCH EASIER with arrays in Perl (mostly).

This object takes a string of values and treats them as an array. If you're using numbers that are more complicated than a double, then this object probably won't work for you. Integers and Longs should be fine. As always, test out your worst case scenario before committing to someone's code.

NOTE: Read the instructions! Early binding is mandatory with this class object.

NOTE: The instructions folder has full explanations and examples for each of these.

Here are the public methods you can access:

<code>objArray.Delimeter</code> => Define what character separates the items in the list. Default is a comma.

<code>objArray.ForceUnique</code> => Turn this on and you will only be able to add unique values to the array.

<code>objArray.ForceValue</code> => Turn this on and you will not be able to add a null value to the array.

<code>objArray.SetArray VBAArray</code> => This is how you populate the array class with an existing VBA array.

<code>objArray.SetArrayFromString strMyStuff</code> => This is how you add a string of items to the array.

<code>objArray.SetArrayFromRange myRangeObject</code> => This is how you add items from a worksheet column.

<code>objArray.Grep([regex pattern])</code> => Gather a list of items from the array but don't change the original array.

<code>objArray.Slice([regex pattern OR index], [limit items])</code> => Gather a list of items from the array AND remove them from the original array.

<code>objArray.Splice([index], [new contet])</code> => Insert one list into another at the specified position.

<code>objArray.Join differentObjArray</code> => COmbine two array objects.

<code>objArray.Push</code> => This will add an item to the back end of the list, thus increasing the size of the list.

<code>objArray.Unshift</code> => This will add an item to the front end of the list, thus increasing the size of the list.

<code>objArray.Pop</code> => This will return a string with the last item in the list and that item will be removed from the list, thus shortening the list.

<code>objArray.Shift</code> => This will return a string with the first item in the list and that item will be removed from the list, thus shortening the list.

<code>objArray.Remove index:=n</code> => This will remove the item at this index position and shorten the list by one item.

<code>objArray.IndexOf("string")</code> => This returns a zero-based index number for the specified items position in the list.

<code>objArray.IndexOfRegex("string")</code> => This returns a zero-based index number for the first item matching the pattern.

<code>objArray.Element(n)</code> => This returns a string with the contents of the element at position N.

<code>objArray.Edit(n) = string</code> => This updates the contents of the element at position N.

<code>objArray.CountElements</code> => This is a 1-based number of the elements in the list. If there are five items in the list, this returns the number 5.

<code>objArray.LastIndex</code> => This returns the index number of the last element in the list. Equivalent to UBound and effectively the same as CountRemaining. I just added this for clarity.

<code>objArray.SortAscending</code> => This sorts the items in the list ascabetically from lowest to highest.

<code>objArray.SortAsNumbers = [boolean]</code> Force the sort to treat the array items as number values instead of string characters.

<code>objArray.Reverse</code> => This reverses the order of the items in the list.

<code>objArray.raw</code> => This is used for debugging. It just returns the array as a string separated by the delimeter.
