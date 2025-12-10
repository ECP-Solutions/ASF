Attribute VB_Name = "testing"
Option Explicit

Public Sub testASFv1()
    Dim engine As ASF
    Dim pidx As Long
    
    Set engine = New ASF
    With engine
        ' Deeply nested arrays //Expect: [10,[20,[30,[40]]]]
'        pidx = .Compile("a = [1,[2,[3,[4]]]]; b = a.map(fun(x) { return x * 10 }); print(b);")
        ' Array of objects + nested arrays //Expect: [{k:2, arr:[11, 21]}, {k:4, arr:[ 31, [41, 51]]}]
'        pidx = .Compile("a = [{ k: 1, arr: [10,20] }, { k: 2, arr: [30,[40,50]] }];" & _
                        "b = a.map(fun(o){return { k: o.k * 2, arr: o.arr.map(fun(x){ return x + 1 })}; });" & _
                        "print(b);")
        ' Closure capture inside map //Expect: [5,10,15]
'        pidx = .Compile("mul = fun(factor){return fun(x){ return x * factor };};" & _
                        "a = [1,2,3]; b = a.map(mul(5));" & _
                        "print(b);")
        ' Map returning nested arrays //Expect: [[1,1],[2,2]]
'        pidx = .Compile("print( [1,2].map(fun(x){ return [x,x] }) );")
        ' Map returning arrays that contain objects that contain arrays
        ' //Expect: [{orig:1, pair:[1, 1], nested:[[1, 2], {v:1}]}, {orig:2, pair:[2, 4], nested:[[2, 3], {v:4}]}]
'        pidx = .Compile("a = [1,2];" & _
                        "b = a.map(fun(n){return {orig: n,pair: [n, n*n],nested: [ [n, n+1], { v: n*n } ]};});" & _
                        "print(b);")
        ' Mapping nested arrays of mixed types //Expect: [3,x,[2,y,[3]]]
'        pidx = .Compile("a = [1,'x',[2,'y',[3]]];" & _
                        "b = a.map(fun(x){if (IsArray(x)) {return x} elseif (IsNumeric(x)) {return x*3} else {return x}};);" & _
                        "print(b);")
        ' Filter simple array //Expect: [ 2, 4 ]
'        pidx = .Compile("a = [1,2,3,4];" & _
                        "b = a.filter(fun(x){ return x % 2 == 0 });" & _
                        "print(b);")
        ' Filter nested arrays //Expect: [ [ 2, 3 ], [ 5 ] ]
'        pidx = .Compile("a=[1,[2,3],4,[5]];" & _
                        "b=a.filter(fun(x){ return IsArray(x) });" & _
                        "print(b);")
        ' Reduce sum with initial //Expect: 10
'        pidx = .Compile("a=[1,2,3,4]; return(a.reduce(fun(acc,x){ return acc + x }, 0));")
        ' Reduce sum with NO initial //Expect: 6
'        pidx = .Compile("a=[1,2,3]; return(a.reduce(fun(acc,x){ return acc + x }));")
        ' Slice with starting point only/ only tail //Expect: [ 20, 30, 40 ]
'        pidx = .Compile("a=[10,20,30,40]; b=a.slice(2); print(b);")
        ' Slice with start and end //Expect: [ 'camel', 'duck' ]
'        pidx = .Compile("a=['ant', 'bison', 'camel', 'duck', 'elephant']; b=a.slice(3,5); print(b);")
        ' Pop and push //Expect: [ 1, 2, 3 ], 4
'        pidx = .Compile("a=[1,2]; a.push(3); a.push(4); x = a.pop(); print(a); print(x);")
        ' Default range //Expect: [ 0, 1, 2 ]
'        pidx = .Compile("print(range(3));")
        ' Custom range //Expect: [ 1, 2 ]
'        pidx = .Compile("print(range(1,3));")
        ' Range with step //Expect: [ 1, 3, 5, 7, 9 ]
'        pidx = .Compile("print(range(1,10,2));")
        ' Flatten full //Expect: [ 1, 2, 3, 4, 5 ]
'        pidx = .Compile("a=[1,[2,3],[4,[5]]]; b = flatten(a); print(b);")
        ' Flatten depth 1 //Expect: [ 1, 2, [ 3 ] ]
        pidx = .Compile("a=[1,[2,[3]]]; b = flatten(a,1); print(b);")
        .Run pidx
    End With
End Sub
