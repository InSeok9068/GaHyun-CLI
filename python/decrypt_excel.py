import msoffcrypto
import io
import shutil
import os

def decrypt_excel_and_save(file_path, password, output_dir):
    try:
        # 엑셀 파일 열기
        with open(file_path, "rb") as file:
            # msoffcrypto-tool 객체 생성
            office_file = msoffcrypto.OfficeFile(file)
            
            # 파일 암호 해제
            office_file.load_key(password=password)
            
            # 암호 해제된 내용을 BytesIO 객체로 추출
            decrypted_stream = io.BytesIO()
            office_file.decrypt(decrypted_stream)

            # 출력 디렉토리와 파일명 설정
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            output_file_path = os.path.join(output_dir, os.path.basename(file_path))
            
            # 암호 해제된 데이터를 지정된 경로에 저장
            with open(output_file_path, "wb") as output_file:
                decrypted_stream.seek(0)  # 스트림의 시작 위치로 이동
                shutil.copyfileobj(decrypted_stream, output_file)
            
            print(f"Decrypted file saved to: {output_file_path}")
            return output_file_path
    except Exception as e:
        print(f"Error: {e}")
        return None

# 사용 예시
if __name__ == "__main__":
    file_path = "./storage/files/sample.xlsm"  # 암호화된 파일 경로
    password = "1111"  # 파일 암호
    output_dir = "./storage/files/temp"  # 해제된 파일을 저장할 디렉토리 경로 
    decrypt_excel_and_save(file_path, password, output_dir)