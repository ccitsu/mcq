// Student Answer Evaluation System - JavaScript
var studentData = [];
var correctAnswers = [];

// File input handling
document.getElementById('fileInput').addEventListener('change', function(e) {
    var file = e.target.files[0];
    if (file) {
        document.getElementById('fileInfo').textContent = 'Selected: ' + file.name;
        document.getElementById('fileInfo').classList.remove('hidden');
        
        var reader = new FileReader();
        reader.onload = function(e) {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, {type: 'array'});
            var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            var jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            studentData = jsonData;
            console.log('Loaded student data:', studentData);
            
            document.getElementById('fileInfo').textContent = 
                'Selected: ' + file.name + ' (' + studentData.length + ' students)';
        };
        reader.readAsArrayBuffer(file);
    }
});

// Auto-uppercase input for answer fields
for (var i = 1; i <= 25; i++) {
    document.getElementById('correct' + i).addEventListener('input', function(e) {
        this.value = this.value.toUpperCase();
    });
}

// Helper functions for answer input
function fillAllAnswers(answer) {
    for (var i = 1; i <= 25; i++) {
        document.getElementById('correct' + i).value = answer;
    }
}

function clearAllAnswers() {
    for (var i = 1; i <= 25; i++) {
        document.getElementById('correct' + i).value = '';
    }
}

// Get correct answers from form
function getCorrectAnswers() {
    var answers = [];
    for (var i = 1; i <= 25; i++) {
        var answer = document.getElementById('correct' + i).value.trim().toUpperCase();
        if (!answer || ['A', 'B', 'C', 'D'].indexOf(answer) === -1) {
            alert('Please enter a valid answer (A, B, C, or D) for Question ' + i);
            return null;
        }
        answers.push(answer);
    }
    return answers;
}

// Calculate grade based on percentage
function calculateGrade(percentage) {
    if (percentage >= 90) return 'A+';
    if (percentage >= 85) return 'A';
    if (percentage >= 80) return 'A-';
    if (percentage >= 75) return 'B+';
    if (percentage >= 70) return 'B';
    if (percentage >= 65) return 'B-';
    if (percentage >= 60) return 'C+';
    if (percentage >= 55) return 'C';
    if (percentage >= 50) return 'C-';
    if (percentage >= 45) return 'D+';
    if (percentage >= 40) return 'D';
    return 'F';
}

// Main evaluation function
function evaluateAnswers() {
    if (studentData.length === 0) {
        alert('Please upload the Excel file first.');
        return;
    }

    correctAnswers = getCorrectAnswers();
    if (!correctAnswers) {
        return;
    }

    // Show loading
    document.getElementById('loadingMsg').classList.add('show');
    document.getElementById('evaluateBtn').disabled = true;

    setTimeout(function() {
        var results = [];
        var totalScore = 0;
        var totalStudents = studentData.length;

        for (var s = 0; s < studentData.length; s++) {
            var student = studentData[s];
            var score = 0;
            var studentAnswers = [];
            var evaluationResults = [];

            // Compare each question
            for (var i = 1; i <= 25; i++) {
                var studentAnswer = (student['Q' + i] || '').toString().trim().toUpperCase();
                var correctAnswer = correctAnswers[i - 1];
                var isCorrect = studentAnswer === correctAnswer;
                
                if (isCorrect) score++;
                
                studentAnswers.push(studentAnswer);
                evaluationResults.push({
                    question: i,
                    studentAnswer: studentAnswer,
                    correctAnswer: correctAnswer,
                    isCorrect: isCorrect
                });
            }

            var percentage = Math.round((score / 25) * 100);
            var grade = calculateGrade(percentage);
            totalScore += score;

            results.push({
                name: student.Name || 'N/A',
                timestamp: student.Timestamp || 'N/A',
                universityId: student.UniversityID || 'N/A',
                email: student.Email || 'N/A',
                score: score,
                percentage: percentage,
                grade: grade,
                answers: studentAnswers,
                evaluation: evaluationResults
            });
        }

        // Display results
        displayResults(results, totalScore, totalStudents);

        // Hide loading
        document.getElementById('loadingMsg').classList.remove('show');
        document.getElementById('evaluateBtn').disabled = false;
    }, 500);
}

// Display results in table
function displayResults(results, totalScore, totalStudents) {
    var avgScore = (totalScore / totalStudents).toFixed(1);
    var avgPercentage = ((totalScore / (totalStudents * 25)) * 100).toFixed(1);
    
    // Calculate pass/fail statistics
    var passCount = 0;
    for (var i = 0; i < results.length; i++) {
        if (results[i].percentage >= 50) passCount++;
    }

    // Display summary statistics
    var summaryHTML = 
        '<div class="bg-blue-100 p-4 rounded-lg text-center">' +
        '<h3 class="font-semibold text-blue-800">Total Students</h3>' +
        '<p class="text-2xl font-bold text-blue-600">' + totalStudents + '</p>' +
        '</div>' +
        '<div class="bg-green-100 p-4 rounded-lg text-center">' +
        '<h3 class="font-semibold text-green-800">Average Score</h3>' +
        '<p class="text-2xl font-bold text-green-600">' + avgScore + '/25</p>' +
        '</div>' +
        '<div class="bg-yellow-100 p-4 rounded-lg text-center">' +
        '<h3 class="font-semibold text-yellow-800">Average Percentage</h3>' +
        '<p class="text-2xl font-bold text-yellow-600">' + avgPercentage + '%</p>' +
        '</div>' +
        '<div class="bg-purple-100 p-4 rounded-lg text-center">' +
        '<h3 class="font-semibold text-purple-800">Pass Rate</h3>' +
        '<p class="text-2xl font-bold text-purple-600">' + passCount + '/' + totalStudents + '</p>' +
        '</div>';
    document.getElementById('summaryStats').innerHTML = summaryHTML;

    // Display detailed results
    var tbody = document.getElementById('resultsTableBody');
    tbody.innerHTML = '';

    for (var r = 0; r < results.length; r++) {
        var result = results[r];
        var row = document.createElement('tr');
        row.className = result.percentage >= 50 ? 'hover:bg-green-50' : 'hover:bg-red-50';
        
        var rowHTML = 
            '<td class="px-4 py-2 border font-medium">' + result.name + '</td>' +
            '<td class="px-4 py-2 border text-sm">' + result.timestamp + '</td>' +
            '<td class="px-4 py-2 border text-center font-bold">' + result.score + '/25</td>' +
            '<td class="px-4 py-2 border text-center font-bold ' + (result.percentage >= 50 ? 'text-green-600' : 'text-red-600') + '">' + result.percentage + '%</td>' +
            '<td class="px-4 py-2 border text-center font-bold">' + result.grade + '</td>';

        // Add individual question results
        for (var i = 0; i < 25; i++) {
            var evaluation = result.evaluation[i];
            var cellClass = evaluation.isCorrect ? 'correct-answer' : 'incorrect-answer';
            var displayText = evaluation.studentAnswer || '-';
            var title = 'Student: ' + (evaluation.studentAnswer || 'No Answer') + ' | Correct: ' + evaluation.correctAnswer;
            
            rowHTML += '<td class="px-2 py-2 border text-center text-sm ' + cellClass + '" title="' + title + '">' + displayText + '</td>';
        }

        row.innerHTML = rowHTML;
        tbody.appendChild(row);
    }

    // Show results section
    document.getElementById('resultsSection').classList.remove('hidden');
    
    // Smooth scroll to results
    document.getElementById('resultsSection').scrollIntoView({ 
        behavior: 'smooth', 
        block: 'start' 
    });

    // Store results globally for export
    window.evaluationResults = results;
}

// Export to CSV function
function exportToCSV() {
    if (!window.evaluationResults) {
        alert('No results to export. Please evaluate answers first.');
        return;
    }

    var csv = 'Student Name,Timestamp,University ID,Email,Score,Percentage,Grade,';
    
    // Add question headers
    for (var i = 1; i <= 25; i++) {
        csv += 'Q' + i + ' Student Answer,Q' + i + ' Correct Answer,Q' + i + ' Result,';
    }
    csv = csv.slice(0, -1) + '\n'; // Remove last comma and add newline

    for (var r = 0; r < window.evaluationResults.length; r++) {
        var result = window.evaluationResults[r];
        var row = '"' + result.name + '","' + result.timestamp + '","' + result.universityId + '","' + result.email + '",' + 
                  result.score + ',' + result.percentage + '%,"' + result.grade + '",';
        
        for (var e = 0; e < result.evaluation.length; e++) {
            var evalResult = result.evaluation[e];
            row += '"' + evalResult.studentAnswer + '","' + evalResult.correctAnswer + '","' + (evalResult.isCorrect ? 'Correct' : 'Incorrect') + '",';
        }
        
        csv += row.slice(0, -1) + '\n'; // Remove last comma and add newline
    }

    // Download CSV
    var blob = new Blob([csv], { type: 'text/csv' });
    var url = window.URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.setAttribute('hidden', '');
    a.setAttribute('href', url);
    a.setAttribute('download', 'evaluation_results_' + new Date().toISOString().split('T')[0] + '.csv');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// Export to Excel function
function exportToExcel() {
    if (!window.evaluationResults) {
        alert('No results to export. Please evaluate answers first.');
        return;
    }

    // Prepare data for Excel
    var wsData = [];
    
    // Headers
    var headers = ['Student Name', 'Timestamp', 'University ID', 'Email', 'Score', 'Percentage', 'Grade'];
    for (var i = 1; i <= 25; i++) {
        headers.push('Q' + i + ' Student', 'Q' + i + ' Correct', 'Q' + i + ' Result');
    }
    wsData.push(headers);

    // Data rows
    for (var r = 0; r < window.evaluationResults.length; r++) {
        var result = window.evaluationResults[r];
        var row = [
            result.name,
            result.timestamp,
            result.universityId,
            result.email,
            result.score,
            result.percentage + '%',
            result.grade
        ];

        for (var e = 0; e < result.evaluation.length; e++) {
            var evalResult = result.evaluation[e];
            row.push(evalResult.studentAnswer, evalResult.correctAnswer, evalResult.isCorrect ? 'Correct' : 'Incorrect');
        }

        wsData.push(row);
    }

    // Create workbook and worksheet
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(wsData);
    
    // Auto-size columns
    var colWidths = [];
    for (var h = 0; h < headers.length; h++) {
        colWidths.push({ wch: 15 });
    }
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, 'Evaluation Results');

    // Save file
    XLSX.writeFile(wb, 'evaluation_results_' + new Date().toISOString().split('T')[0] + '.xlsx');
}

// Print results function
function printResults() {
    if (!window.evaluationResults) {
        alert('No results to print. Please evaluate answers first.');
        return;
    }

    var printWindow = window.open('', '_blank');
    
    printWindow.document.write('<html><head><title>Evaluation Results</title>');
    printWindow.document.write('<style>');
    printWindow.document.write('body { font-family: Arial, sans-serif; margin: 20px; }');
    printWindow.document.write('table { border-collapse: collapse; width: 100%; }');
    printWindow.document.write('th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }');
    printWindow.document.write('th { background-color: #f2f2f2; }');
    printWindow.document.write('.correct-answer { background-color: #22c55e !important; color: white; }');
    printWindow.document.write('.incorrect-answer { background-color: #ef4444 !important; color: white; }');
    printWindow.document.write('@media print { body { margin: 0; } table { font-size: 10px; } }');
    printWindow.document.write('</style></head><body>');
    printWindow.document.write('<h1>Student Answer Evaluation Results</h1>');
    printWindow.document.write('<p>Generated on: ' + new Date().toLocaleString() + '</p>');
    
    var printContent = document.getElementById('resultsSection').innerHTML;
    printWindow.document.write(printContent);
    
    printWindow.document.write('</body></html>');
    printWindow.document.close();
    printWindow.print();
}

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    // Disable evaluate button initially
    document.getElementById('evaluateBtn').disabled = true;
    
    // Enable evaluate button when both file and answers are ready
    function checkReadiness() {
        var hasFile = studentData.length > 0;
        var hasAnswers = getCorrectAnswers() !== null;
        document.getElementById('evaluateBtn').disabled = !(hasFile && hasAnswers);
    }

    // Check readiness when file changes
    document.getElementById('fileInput').addEventListener('change', function() {
        setTimeout(checkReadiness, 100);
    });

    // Check readiness when answers change
    for (var i = 1; i <= 25; i++) {
        document.getElementById('correct' + i).addEventListener('input', checkReadiness);
    }
});