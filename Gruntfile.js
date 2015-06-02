module.exports = function(grunt) {
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json')
    });
    grunt.initConfig({
        coffee: {
            compile:{
                files: [{
                    expand: true,
                    cwd: 'src/',
                    src: ['**/*.coffee'],
                    dest: 'build/',
                    ext: '.js'
                },
                {
                    expand: true,
                    cwd: './',
                    src: ['*.coffee'],
                    dest: './',
                    ext: '.js'
                }]
            }
        }
    });
    grunt.loadNpmTasks('grunt-contrib-coffee');
    grunt.registerTask('default', ['coffee']);
};