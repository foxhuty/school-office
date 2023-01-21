import datetime
from flask import Flask, render_template, request, send_from_directory, flash
from scores_source.scores_process import ScoreAnalysis
from scores_source.scores_process import JuniorExam
from scores_source.scores_process import ExamRoom
from scores_source.scores_process import ExamInvigilators
from scores_source.scores_process import SplitClass
from scores_source.scores_process import CatalogueCourses
from scores_source.scores_process import GetInfoFromId
import os
from concurrent.futures import ThreadPoolExecutor

pool = ThreadPoolExecutor(50)

app = Flask(__name__)
app.config['SECRET_KEY']='lsajfl;dsjfo[ujsdfljsda'

root_dir = r'D:\成绩统计结果'


def del_files(file_dir):
    n = 0
    if len(os.listdir(file_dir)) >= 10:
        for f in os.listdir(file_dir):
            os.remove(file_dir + os.sep + f)
        print(f'files have been deleted')
        n += 1
    print(f'第{n}次清理完毕')


def get_date_weekday():
    date = datetime.datetime.now().strftime('%Y年%m月%d日')
    week = datetime.datetime.weekday(datetime.datetime.now())
    week_name = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    week_day = week_name[week]
    date_week = f'{date} {week_day}'
    return date_week


def data_update():
    date = datetime.datetime.today()
    date = date.strftime('%Y年%m月%d日 %H点%M分%S秒')
    return date


@app.route('/')
def instructions():
    return render_template('instructions.html')


@app.route('/index/', methods=['GET', 'POST'])
def process_by_total():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    arts_scores = request.form.get('arts_scores')
    science_scores = request.form.get('science_scores')
    if request.method == 'GET':
        return render_template('index.html',
                               update_time=update_time,
                               weekday=weekday)
    else:
        try:
            exam = ScoreAnalysis(path_file)
            ScoreAnalysis.arts_scores = exam.get_goodscores_arts(float(arts_scores))
            ScoreAnalysis.science_scores = exam.get_goodscores_science(float(science_scores))
            df_arts, df_science = exam.get_av()
            df_arts.reset_index(inplace=True)
            df_arts = df_arts.to_html(index=False, classes=['table', 'table-bordered', 'table-striped'], na_rep='')
            df_science.reset_index(inplace=True)
            df_science = df_science.to_html(index=False, classes=['table', 'table-bordered', 'table-striped'],
                                            na_rep='')
            arts_grade = exam.subject_grade_arts()
            science_grade = exam.subject_grade_science()
            df_html = arts_grade.to_html(index=True, classes=['table', 'table-bordered', 'table-striped'], na_rep='')
            df_html_science = science_grade.to_html(index=True, classes=['table', 'table-bordered', 'table-striped'],
                                                    na_rep='')
            result_processed = exam.arts_science_combined_school(goodtotal_arts=float(arts_scores),
                                                                 goodtotal_science=float(science_scores))
            return render_template('index.html',
                                   weekday=weekday,
                                   update_time=update_time,
                                   df_html=df_html,
                                   df_science_html=df_html_science,
                                   av_html=df_arts,
                                   av_science_html=df_science,
                                   path_file=path_file,
                                   arts_scores=arts_scores,
                                   science_scores=science_scores,
                                   result_processed=result_processed)
        except:
            return render_template('index.html',
                                   update_time=update_time,
                                   weekday=weekday)


@app.route('/index/<filename>')
def download_by_total(filename):
    return send_from_directory(root_dir, filename)


@app.route('/scores/<filename>')
def download_by_goodscores(filename):
    return send_from_directory(root_dir, filename)


@app.route('/scores/', methods=['GET', 'POST'])
def process_by_good_scores():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    arts_scores = request.form.get('arts_scores')
    science_scores = request.form.get('science_scores')
    if request.method == 'GET':
        return render_template('scores.html',
                               update_time=update_time, weekday=weekday)
    else:
        try:
            exam = ScoreAnalysis(path_file)
            ScoreAnalysis.arts_scores = eval(arts_scores)
            ScoreAnalysis.science_scores = eval(science_scores)
            df_arts, df_science = exam.get_av()
            df_arts.reset_index(inplace=True)
            df_arts = df_arts.to_html(index=False, classes=['table', 'table-bordered'], na_rep='')
            df_science.reset_index(inplace=True)
            df_science = df_science.to_html(index=False, classes=['table', 'table-bordered'], na_rep='')
            arts_grade = exam.subject_grade_arts()
            science_grade = exam.subject_grade_science()
            arts_grade = arts_grade.to_html(index=True, classes=['table', 'table-bordered'], na_rep='')
            science_grade = science_grade.to_html(classes=['table', 'table-bordered'], na_rep='')
            exam_processed = exam.arts_science_combined()

            return render_template('scores.html',
                                   weekday=weekday,
                                   update_time=update_time,
                                   df_html=arts_grade,
                                   df_science_html=science_grade,
                                   av_html=df_arts,
                                   av_science_html=df_science,
                                   path_file=path_file,
                                   arts_scores=arts_scores,
                                   science_scores=science_scores,
                                   exam_processed=exam_processed)
        except:
            return render_template('scores.html',
                                   update_time=update_time, weekday=weekday)


@app.route('/exams/', methods=['GET', 'POST'])
def exams():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    room_number = request.form.get('room_number')
    try:
        exam_rooms = ExamRoom(path_file)
        ExamRoom.room_num = int(room_number)
        exams_info = exam_rooms.exam_room_choice()
        return render_template('exams.html',
                               weekday=weekday,
                               update_time=update_time,
                               path_file=path_file,
                               room_number=room_number,
                               exams_info=exams_info)
    except:
        return render_template('exams.html',
                               update_time=update_time, weekday=weekday)


@app.route('/exams/<filename>')
def download_by_exams(filename):
    return send_from_directory(root_dir, filename)


@app.route('/split_class/', methods=['GET', 'POST'])
def divide_into_class():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    try:
        divided = SplitClass(path_file)
        divided_class = divided.split_into_class()
        return render_template('divided_class.html',
                               weekday=weekday,
                               update_time=update_time,
                               path_file=path_file,
                               divided_class=divided_class
                               )
    except:
        return render_template('divided_class.html',
                               update_time=update_time,
                               weekday=weekday)


@app.route('/split_class/<filename>')
def download_divided_class(filename):
    return send_from_directory(root_dir, filename)


@app.route('/invigilation/', methods=['GET', 'POST'])
def invigilation():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    exam_numbers = request.form.get('exam_numbers')
    room_numbers = request.form.get('room_numbers')
    try:
        exam = ExamInvigilators(path_file)
        ExamInvigilators.room_numbers = int(room_numbers)
        ExamInvigilators.exam_numbers = int(exam_numbers)
        invigilator_table = exam.invigilation_table()
        invigilation_html = invigilator_table.to_html(index=True, classes=['table', 'table-bordered'], na_rep='')
        invigilators = exam.exam_teachers()
        return render_template('invigilation.html',
                               invigilation_html=invigilation_html,
                               invigilators=invigilators,
                               path_file=path_file,
                               exam_numbers=exam_numbers,
                               room_numbers=room_numbers,
                               update_time=update_time,
                               weekday=weekday)
    except:
        return render_template('invigilation.html',
                               update_time=update_time,
                               weekday=weekday)


@app.route('/junior/', methods=['GET', 'POST'])
def junior():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    good_scores = request.args.get('good_scores')
    try:
        junior_exam = JuniorExam(path_file)
        junior_scores = junior_exam.junior_scores()
        junior_av = junior_exam.get_av()
        junior_scores_html = junior_scores.to_html(index=False, classes=['table', 'table-bordered'], na_rep='')
        junior_av_html = junior_av.to_html(index=False, classes=['table', 'table-bordered'], na_rep='')
        junior_goodscores = junior_exam.goodscore_process(float(good_scores))
        goodscore_html = junior_goodscores.to_html(index=True, classes=['table', 'table-bordered'], na_rep='')
        final_export = junior_exam.concat_results(float(good_scores))
        return render_template('junior.html',
                               weekday=weekday,
                               update_time=update_time,
                               junior_scores_html=junior_scores_html,
                               junior_av_html=junior_av_html,
                               path_file=path_file,
                               goodscore_html=goodscore_html,
                               final_export=final_export)
    except:
        return render_template('junior.html',
                               update_time=update_time, weekday=weekday)


@app.route('/junior/<filename>')
def junior_download(filename):
    return send_from_directory(root_dir, filename)


@app.route('/invigilation/<filename>')
def invigilation_download(filename):
    return send_from_directory(root_dir, filename)


@app.route('/catalog/', methods=['GET', 'POST'])
def catalog_courses():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    try:
        catalog = CatalogueCourses(path_file)
        catalog_table = catalog.statistics_for_courses()
        catalog_html = catalog_table.to_html(index=False, classes=['table', 'table-bordered', 'table-striped'],
                                             na_rep='')
        catalog_exported = catalog.split_by_subject()
        return render_template('catalogue_courses.html',
                               catalog_html=catalog_html,
                               update_time=update_time,
                               weekday=weekday,
                               path_file=path_file,
                               catalog_exported=catalog_exported)
    except:
        return render_template('catalogue_courses.html', update_time=update_time, weekday=weekday)


@app.route('/catalog/<filename>')
def catalog_download(filename):
    return send_from_directory(root_dir, filename)


@app.route('/get_info/', methods=['GET', 'POST'])
def get_info():
    weekday = get_date_weekday()
    update_time = data_update()
    path_file = request.files.get('path_file')
    try:
        id_info = GetInfoFromId(path_file)
        id_matched, bad_df, id_bad, show_data = id_info.get_sex_birth_age()
        bad_df_html = bad_df.to_html(index=False,
                                     classes=['table', 'table-bordered', 'table-striped'],
                                     formatters={
                                         'text-align': 'center',
                                         'color': 'red'
                                     },
                                     na_rep='')
        show_data_html = show_data.to_html(index=False, classes=['table', 'table-bordered', 'table-striped'],
                                           formatters={
                                               'text-align': 'center'
                                           })
        return render_template('get_info_from_id.html',
                               bad_df=bad_df,
                               id_matched=id_matched,
                               bad_df_html=bad_df_html,
                               id_bad=id_bad,
                               update_time=update_time,
                               weekday=weekday,
                               path_file=path_file,
                               show_data_html=show_data_html)
    except:
        flash('请输入正确的文件名或文件格式！！！')
        return render_template('get_info_from_id.html', update_time=update_time, weekday=weekday)


@app.route('/get_info/<filename>')
def id_info_download(filename):
    return send_from_directory(root_dir, filename)


if __name__ == '__main__':
    pool.submit(app.run)
    # app.run()
