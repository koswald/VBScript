###### The `key` folder

# ReadMe.md

To continue, you must either [generate a strong-name key pair]
(recommended), or else [opt out].  

When you specify a key file, the compiler uses it to 
digitally sign the output file.

[generate a strong-name key pair]: #to-generate-a-key-pair
[opt out]: #to-opt-out

> Note: It is recommended that you keep _your_ keyfile 
> out of the project folders so that it doesn't
> get checked out by accident. 

### To generate a key pair

Visual Studio is required to generate a strong-name 
key pair.

1. You can use [generate-key-pair.vbs] to 
   generate a key pair, after editing it to 
   give a unique name to the file and change the location.  

   Alternatively, open the Visual Studio 
   Developer Command Prompt and type

   `sn -k <keyfile>`  

    where `<keyfile>` specifies the location and 
    name of the keyfile to be generated. Typically 
    the `.snk` extension is used.

2. Then edit [_common.rsp] to indicate the name 
   and location of your key file.

### To opt out

Edit [_common.rsp]. Comment out or remove the 
line with `/key` and enable the line 
with `/delaysign` by removing the `#`.

[_common.rsp]: ../rsp/_common.rsp

[generate-key-pair.vbs]: generate-key-pair.vbs

<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />

