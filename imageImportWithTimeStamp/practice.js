const str = 'I like cats';
const before = 'cats';
const after = 'dogs';

r = ((str, before, after) => str.replace(before, after))(str, before, after);

console.log(r); // I like dogs