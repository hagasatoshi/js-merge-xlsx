(function() {
  var SampleOuter, sample;

  SampleOuter = (function() {
    var shared_strings;

    function SampleOuter() {
      this.message = 'This is Test';
      this.shared_strings = new shared_strings();
      console.log(this.shared_strings.counter);
      this.shared_strings.countup();
      console.log(this.shared_strings.counter);
    }

    shared_strings = (function() {
      function shared_strings() {
        this.counter = 0;
      }

      shared_strings.prototype.countup = function() {
        return this.counter++;
      };

      return shared_strings;

    })();

    return SampleOuter;

  })();

  sample = new SampleOuter();

}).call(this);
