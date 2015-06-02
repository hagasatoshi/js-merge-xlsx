class SampleOuter
  constructor: ()->
    @message = 'This is Test'
    @shared_strings = new shared_strings()
    console.log @shared_strings.counter
    @shared_strings.countup()
    console.log @shared_strings.counter
    
  class shared_strings
    constructor: ()->
      @counter = 0
    countup: ()->
      @counter++
      
sample = new SampleOuter()