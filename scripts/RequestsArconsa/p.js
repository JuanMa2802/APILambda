let lista = [1,2,3,4]

let clone = Array.from(lista.filter(el=>{if (el==1) { } else {
    return el
  }}))

console.log(clone)