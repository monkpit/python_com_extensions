## Python COM Object Extensions for PyWin32

### What is it?

If you've ever tried to add any properties or methods to a pywin32 COM object, you may have run into an issue:

### You can't.

Not easily, anyways. Except now you can!

### Requirements

* PyWin32

### The problem

Let's say you've got a `MyCom.Application` object, and the `IList` interface contains a collection.
But for some reason, pywin32 doesn't feel like being helpful...
You try to use a list comprehension and you're getting nowhere fast:

```python
    import win32com.client
    mycom_app = win32com.client.Dispatch('MyCom.Application')
    list_of_things = [o.Name for o in mycom_app.collection if 'test' in o.Name]
```

And, BAM. Your code explodes into a million pieces because you tried to call `__iter__`.

`TypeError: This object does not support enumeration`.

WTF, pywin32? Why didn't you implement `__iter__`?

No matter... we will handle this ourselves.


### The solution

Just define a simple class declaring what methods you'd like to add or override.

```python
    from com_extensions import ComExtension
    
    class MyComApplicationExtensions(ComExtension):
        progid = 'MyCom.Application'
        extends = 'IList'
        
        def __iter__(self):
            current = 1
            while current <= self.Count:
                yield self.Item(current)
                current += 1
```

Then, before you create your COM object, sprinkle in a little magic:

```python
    import win32com.client
    
    def main():
        MyComApplicationExtensions.register()
        mycom_app = win32com.client.Dispatch('MyCom.Application')
        list_of_things = [o.Name for o in mycom_app.collection if 'test' in o.Name]
        print(list_of_things)
        
    # prints: ['test1', 'test2', 'this_is_a_test'] - or whatever.
```
