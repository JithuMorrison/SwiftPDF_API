from flask import Flask, jsonify, request 
import joblib

# creating a Flask app 
app = Flask(__name__) 
  
@app.route('/', methods=['GET', 'POST']) 
def home(): 
    data = request.get_json()
    # Do something with the data
    response_data = {
        'received_json': data
    }

    # Return a JSON response
    return jsonify(response_data)
  
@app.route('/home', methods=['GET','POST']) 
def disp(): 
    data = request.get_json()
    print(data["text"])
    # Load the saved model and vectorizer
    model = joblib.load('trauma_model.pkl')
    vectorizer = joblib.load('vectorizer.pkl')

    # Function to make predictions on new data
    def predict_trauma(text):
        text_tfidf = vectorizer.transform([text])
        prediction = model.predict(text_tfidf)
        return prediction[0]

    # Example usage of the prediction function
    new_text = data["text"]
    predicted_type = predict_trauma(new_text)
    print(f"The predicted trauma type is: {predicted_type}")
    response_data = {
        'received_json': predicted_type
    }
    return jsonify(response_data) 

@app.route('/trauma', methods=['POST'])
def predict_trauma_route():
    data = request.get_json()
    print(data["text"])
    loaded_pipeline = joblib.load('trained_model.joblib')
    # Example usage
    new_text = data["text"]  # Replace with the text you want to classify
    predicted_label = loaded_pipeline.predict([new_text])
    print(f"Predicted label: {predicted_label[0]}")
    response_data = {
        'predict': str(predicted_label[0])
    }
    
    return jsonify(response_data)
  
if __name__ == '__main__': 
    app.run() 
