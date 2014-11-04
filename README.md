AutoItObject v1.2.8.3
=====================
### Bringing Objects to AutoIt

#### For AutoIt Version 3.3

[![Build status](https://ci.appveyor.com/api/projects/status/kvmrr2jngea9fsni)](https://ci.appveyor.com/project/KyleChamberlin/autoitobject)

This is a fork created for the AutoItMicro project. There have been no changes to the code, it has been integrated with Appveyor CI and made into a submodule for use with Micro Unit Testing Framework.

Using AutoItObject
------------------

To use AutoItObject you simply need to add it as a submodule to the project you'd like to test.

    cd <your project's root>
    git submodule add git://github.com/AutoItMicro/AutoItObject.git
    git submodule update --init --recursive

now if you add `AutoItObject\AutoItObject.au3` as an include in your test script you are good to go. 
It is recommended to use `#include-once` in your script to avoid collisions. like this:

```AutoIt
#include-once AutoItObject\AutoItObject.au3
```

Authors
-------
- Andreas Karlsson (monoceres)
- Dragana R. (trancexx)
- Dave Bakker (Kip)
- Andreas Bosch (progandy, Prog@ndy)

Appveyor integration by: @KyleChamberlin

License
-------

Artistic License 2.0, see LICENSE
