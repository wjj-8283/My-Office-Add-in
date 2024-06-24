# @dbpiper/timer

[![Build Status](https://travis-ci.com/dbpiper/timer.svg?branch=master)](https://travis-ci.com/dbpiper/timer)
[![npm version](http://img.shields.io/npm/v/@dbpiper/timer.svg?style=flat)](https://npmjs.org/package/@dbpiper/timer 'View this project on npm')
[![MIT license](http://img.shields.io/badge/license-MIT-brightgreen.svg)](http://opensource.org/licenses/MIT)

Simple module to time your code JavaScript or TypeScript code, to see how long
it took to run.

## Installation Instructions

```sh
npm install @dbpiper/timer
```

## Usage Instructions

```ts
import Timer, { FormattedDuration } from '@dbpiper/timer';

// this starts the timer running
const timer = new Timer();

// you may also start the timer like this
// if you wanted to create the timer instance first and then do some
// setup before starting the timer for instance
timer.start();

// do some stuff here

// formattedDuration has the amount of time that passed between
// the creation and stop times
const formattedDuration = timer.stop();

// you can simply convert the result formattedDuration to a string either
// implicitly or explicitly
// for example:
console.log(formattedDuration);
```

## API

### Timer

The actual instance object that is used to access and use the timer.
The timer will be started automatically when an instance is created.

### timer.start()

An alternative, more explicit, way to start the timer. This is automatically
called by the constructor, however if you want to do some tasks after creating
the Timer instance, but before starting the measurement then this can be used.

### timer.stop()

return type: `FormattedDuration`

This stops the timer, and returns the amount of time that passed as a
FormattedDuration instance object.

---

### FormattedDuration

The duration that passed, it is meant for easy displaying in scripts as the
main purpose of the module is to measure how much time a script took to run.
Thus, you would start the timer right before running it and then stop it when
done, printing the result to the screen or a log file or a message to an
external service.

### formattedDuration.hours

return type: `number`

The amount of hours in the duration, note that since the largest unit supported
is hours everything larger will be here. So for example 2 days would be here
as 48 hours, since there is no "days" unit. The reason for this being the
largest unit supported is the fact that this module is intended to be used for
timing code.

If your code is running for days or weeks, this probably isn't the right module
to use, rather it is intended to be used for modules that run on the order
of minutes or maybe a few hours.

### formattedDuration.minutes

return type: `number`

The amount of minutes in the duration.

### formattedDuration.seconds

return type: `number`

The amount of seconds in the duration.

### formattedDuration.milliseconds

return type: `number`

The smallest unit that _@dbpiper/timer_ supports.

### formattedDuration.toString()

Converts the duration to a human-readable string, it is usually used
implicitly like this:

```ts
console.log(formattedDuration);
// or this

const durationString = `${formattedDuration}`;
```

It will print each of the usual units described above, starting from
the largest one that is non-zero. Each unit will be joined with a comma and
space, except for the last one which will have "and" instead.

For example:

`30 hours, 31 minutes and 26.4 seconds`

#### No Milliseconds in output

You can observe from this example that there are no milliseconds in the
string! This is because I feel that adding them would not help to serve
the purpose of this module, which is to measure the amount of time that
code took to run and report it to a log or script.

This alterative _with milliseconds_:

`30 hours, 31 minutes, 26 seconds and 400 milliseconds`

Is about 40% longer, and contributes almost no useful information. People
still look at something like: `400 milliseconds` and say "hmm that's about
half a second". In other words, leaving the milliseconds out and making them
part of the seconds is not only shorter in your log, its actually easier for
you to read!

#### Some Milliseconds in output

You might be thinking, "I noticed earlier that I can access the milliseconds
explicitly with formattedDuration.milliseconds, why would you put that in there
if it isn't used?". The answer to this question is simple: it **is** used, but
only if the duration is exceptionally short. Specifically, if it is less than
a second.

For example a duration of 0.4 seconds or 400 milliseconds would **not** be
printed as `0.4 seconds`. This has the reverse [problem from above](#no-milliseconds-in-output).
That is, it is _too_ terse making it harder to read and understand. It is much
easier to understand `400 milliseconds` rather than `0.4 seconds`, simply because
when you get to such a short time-frame it makes more sense to think in terms of
explicit milliseconds.

This is likely not going to come up too often, unless you are timing very short
segments of code. However, just note that in these cases you **will** get
milliseconds. In other words, you have them when you need them!

## License

[MIT](https://github.com/dbpiper/timer/blob/master/LICENSE) Copyright
© [David Piper](https://github.com/dbpiper)
