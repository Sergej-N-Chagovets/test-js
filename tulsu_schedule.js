//region Init
let min_date
let max_date
let today = new Date()
let search_field
let search_value
const is_safari = /^((?!chrome|android).)*safari/i.test(navigator.userAgent)

$('.search').autocomplete({
	minLength: 3,
	source: (request, response) => {
		$.ajax({
			url: 'queries/GetDictionaries.php',
			type: 'POST',
			data: {
				search_value: request.term,
			},
			dataType: 'JSON',
			success: response,
		})
	},
	select: (event, ui) => {
		$('.search').val(ui.item['value'])
		Search()
	}
})

const searchParams = new URLSearchParams(location.search)
if (searchParams.has('group')) {
	searchParams.set('search', searchParams.get('group'))
	searchParams.delete('group')
	const new_url = location.pathname + '?' + searchParams.toString()
	history.replaceState(null, null, new_url)
} else if (searchParams.has('teacher')) {
	searchParams.set('search', searchParams.get('teacher'))
	searchParams.delete('teacher')
	const new_url = location.pathname + '?' + searchParams.toString()
	history.replaceState(null, null, new_url)
} else if (searchParams.has('aud')) {
	searchParams.set('search', searchParams.get('aud'))
	searchParams.delete('aud')
	const new_url = location.pathname + '?' + searchParams.toString()
	history.replaceState(null, null, new_url)
}

if (searchParams.get('display') === 'list' || searchParams.get('display') === 'table') {
	localStorage.setItem('display-mode', searchParams.get('display'))
	searchParams.delete('display')
	const new_url = location.pathname + (searchParams.toString() ? '?' + searchParams.toString() : '')
	history.replaceState(null, null, new_url)
}

if (!localStorage.getItem('display-mode')) {
	localStorage.setItem('display-mode', 'table')
}

if (localStorage.getItem('display-mode') === 'table') {
	$('.display-mode').addClass('display-mode__table')
	$('.header, .settings-container').css('position', 'static')
} else {
	$('.display-mode').addClass('display-mode__list')
	$('.header, .settings-container').css('position', 'sticky')
}

if (searchParams.has('search')) {
	$('.search').val(searchParams.get('search'))
	Search()
}

//endregion

//region Functions
function Search() {
	search_value = $('.search').val()
	if (search_value && $('.zv-load-css').css('visibility') === 'hidden') {
		$('.zv-load-css').css('visibility', 'visible')
		$('.calendar-container').hide()
		$('.print').hide()
		$('.table-container').slideUp()
		$('.list-container').slideUp()
		$('.message-not-found').slideUp()

		searchParams.set('search', search_value)
		const new_url = location.pathname + '?' + searchParams.toString()
		if (new_url !== location.pathname + location.search) {
			history.pushState(null, null, new_url)
		}

		if (localStorage.getItem('display-mode') === 'table') {
			AddDates()
		} else {
			AddListDates()
		}

		ShowCalendar()
	}
}

function AddDates() {
	$.ajax({
		url: 'queries/GetDates.php',
		type: 'POST',
		data: {
			search_value: search_value,
		},
		dataType: 'JSON',
		success: data => {
			if (data['error']) {
				$('.zv-load-css').css('visibility', 'hidden')
				$('.message-not-found').html(data['error'])
				$('.message-not-found').slideDown()
			} else {
				search_field = data['SEARCH_FIELD']

				min_date = new Date(data['MIN_DATE'])
				max_date = new Date(data['MAX_DATE'])

				let html = '<tr><th class="schedule-time"></th>'
				for (let this_date = new Date(min_date); this_date <= max_date; this_date.setDate(this_date.getDate() + 1)) {
					const this_date_str = this_date.toLocaleDateString('ru-RU', {
						weekday: 'short',
						day: '2-digit',
						month: '2-digit',
					})
					const this_date_str_short = this_date.toLocaleDateString('ru-RU')
					const is_today = this_date_str_short === today.toLocaleDateString('ru-RU')
					const is_monday = this_date.getDay() === 1
					html += `<th class="schedule-date ${is_today ? 'today' : ''} ${is_monday ? 'monday' : ''}" data-date="${this_date_str_short}"><div>${this_date_str}</div></th>`
				}
				html += '</tr>'

				$('.schedule').html(html)

				AddTimeGroups()
			}
		},
		error: error => {
			$('.zv-load-css').css('visibility', 'hidden')
			$('.message-not-found').html(error.responseText)
			$('.message-not-found').slideDown()
		},
	})
}

function AddTimeGroups() {
	$.ajax({
		url: 'queries/GetTimeGroups.php',
		type: 'POST',
		data: {
			search_field: search_field,
			search_value: search_value,
		},
		dataType: 'JSON',
		success: data => {
			data.forEach(item => {
				let html = `<tr><td class="schedule-time">${item['TIME_START']}<br>-<br>${item['TIME_END']}</td>`
				for (let this_date = new Date(min_date); this_date <= max_date; this_date.setDate(this_date.getDate() + 1)) {
					const this_date_str = this_date.toLocaleDateString('ru-RU')
					const is_today = this_date_str === today.toLocaleDateString('ru-RU')
					html += `
                        <td class="${is_today ? 'today' : ''}">
                            <div class="schedule-lessons" data-date="${this_date_str}" data-time="${item['TIME_START']} - ${item['TIME_END']}"></div>
                        </td>`
				}
				html += '</tr>'

				$('.schedule').append(html)
			})

			AddLessons()
		},
		error: error => {
			$('.zv-load-css').css('visibility', 'hidden')
			$('.message-not-found').html(error.responseText)
			$('.message-not-found').slideDown()
		},
	})
}

function AddLessons() {
	$.ajax({
		url: 'queries/GetSchedule.php',
		type: 'POST',
		data: {
			search_field: search_field,
			search_value: search_value,
		},
		dataType: 'JSON',
		success: data => {
			data.forEach(item => {
				const subgroup = search_field === 'GROUP_P' ? item['GROUPS'][0]['PRIM'] : null
				let html = `
                    <div class="schedule-lesson ${item['CLASS']}">
                        <div class="schedule-lesson-info">${item['DISCIP']}</div>
                        <div class="schedule-lesson-info">(${item['KOW']})${subgroup ? ` (${subgroup})` : ''}</div>
						<div class="schedule-lesson-info">${item['TIME_Z']}</div>`
				if (search_field !== 'AUD' && item['AUD']) {
					html += `<a class="schedule-lesson-info" href="${'/schedule/?search=' + encodeURIComponent(item['AUD'])}">${item['AUD']}</a>`
				}
				if (search_field !== 'PREP' && item['PREP']) {
					html += `<a class="schedule-lesson-info" href="${'/schedule/?search=' + encodeURIComponent(item['PREP'])}">${item['PREP']}</a>`
				}
				if (search_field !== 'GROUP_P' && item['GROUPS']) {
					item['GROUPS'].forEach(group => {
						const group_name = group['GROUP_P'] + (group['PRIM'] ? ` (${group['PRIM']})` : '')
						html += `<a class="schedule-lesson-info group-link" href="${'/schedule/?search=' + encodeURIComponent(group['GROUP_P'])}">${group_name}</a>`
					})
				}
				html += '<div class="more-groups"></div></div>'

				$(`.schedule-lessons[data-date="${item['DATE_Z']}"][data-time="${item['TIME_Z']}"]`).append(html)
			})

			$('.zv-load-css').css('visibility', 'hidden')
			$('.table-container').slideDown('fast', () => {
				HideMultiGroups()
				ScrollSchedule()
			})
			$('.print').show('slow')
		},
		error: error => {
			$('.zv-load-css').css('visibility', 'hidden')
			$('.message-not-found').html(error.responseText)
			$('.message-not-found').slideDown()
		},
	})
}

function AddListDates() {
	$.ajax({
		url: 'queries/GetDates.php',
		type: 'POST',
		data: {
			search_value: search_value,
		},
		dataType: 'JSON',
		success: data => {
			if (data['error']) {
				$('.zv-load-css').css('visibility', 'hidden')
				$('.message-not-found').html(data['error'])
				$('.message-not-found').slideDown()
			} else {
				search_field = data['SEARCH_FIELD']

				min_date = new Date(data['MIN_DATE'])
				max_date = new Date(data['MAX_DATE'])

				$('.list-container').html('')
				for (let this_date = new Date(min_date); this_date <= max_date; this_date.setDate(this_date.getDate() + 1)) {
					let this_date_str = this_date.toLocaleDateString('ru-RU', {
						weekday: 'long',
						day: '2-digit',
						month: '2-digit',
					})
					this_date_str = this_date_str.slice(0, 1).toUpperCase() + this_date_str.slice(1)
					const this_date_str_short = this_date.toLocaleDateString('ru-RU')
					const is_today = this_date_str_short === today.toLocaleDateString('ru-RU')
					$('.list-container').append(`
						<div class="list-day content-wrapper ${is_today ? 'today' : ''}" data-date="${this_date_str_short}" style="display: none">
							<div class="list-date">${this_date_str}</div>
						</div>`
					)
				}

				AddListLessons()
			}
		},
		error: error => {
			$('.zv-load-css').css('visibility', 'hidden')
			$('.message-not-found').html(error.responseText)
			$('.message-not-found').slideDown()
		},
	})
}

function AddListLessons() {
	$.ajax({
		url: 'queries/GetSchedule.php',
		type: 'POST',
		data: {
			search_field: search_field,
			search_value: search_value,
		},
		dataType: 'JSON',
		success: data => {
			data.forEach(item => {
				const subgroup = search_field === 'GROUP_P' ? item['GROUPS'][0]['PRIM'] : null
				let html = `
					<div class="list-row">
						<div class="list-time">${item['TIME_Z']}</div>
                	    <div class="list-lesson ${item['CLASS']}">
                	        <div class="list-lesson-info">${item['DISCIP']} (${item['KOW']})${subgroup ? ` (${subgroup})` : ''}</div>`
				if (search_field !== 'AUD' && item['AUD']) {
					html += `<a class="list-lesson-info" href="${'/schedule/?search=' + encodeURIComponent(item['AUD'])}">${item['AUD']}</a>`
				}
				if (search_field !== 'PREP' && item['PREP']) {
					html += `<a class="list-lesson-info" href="${'/schedule/?search=' + encodeURIComponent(item['PREP'])}">${item['PREP']}</a>`
				}
				if (search_field !== 'GROUP_P' && item['GROUPS']) {
					html += '<div class="list-lesson-groups">'
					item['GROUPS'].forEach(group => {
						const group_name = group['GROUP_P'] + (group['PRIM'] ? ` (${group['PRIM']})` : '')
						html += `<a class="list-lesson-info list-group-link" href="${'/schedule/?search=' + encodeURIComponent(group['GROUP_P'])}">${group_name}</a>`
					})
					html += '</div>'
				}
				html += '</div></div>'

				$(`.list-day[data-date="${item['DATE_Z']}"]`).append(html)
				$(`.list-day[data-date="${item['DATE_Z']}"]`).show()
			})

			$('.zv-load-css').css('visibility', 'hidden')
			$('.list-container').slideDown('fast', ScrollList)
			$('.print').show('slow')
		},
		error: error => {
			$('.zv-load-css').css('visibility', 'hidden')
			$('.message-not-found').html(error.responseText)
			$('.message-not-found').slideDown()
		},
	})
}

function ShowCalendar() {
	$.ajax({
		url: 'queries/GetCalendar.php',
		data: {
			search_value: search_value,
		},
		dataType: 'JSON',
		success: data => {
			$('.calendar-list').html('')
			data.forEach(item => {
				$('.calendar-list').append(`<div>${item['BEGIN_DATE']} - ${item['END_DATE']} ${item['VID']}</div>`)
			})
			if (data.length > 0) {
				$('.calendar-container').show('slow')
			}
		},
	})
}

function GetHiddenHeight(item) {
	return item.find('.group-link').eq(3).offset().top + item.find('.group-link').eq(3).outerHeight() + parseInt(item.css('padding-bottom')) - item.offset().top
}

function GetFullHeight(item) {
	return item.find('.group-link').last().offset().top + item.find('.group-link').last().outerHeight() + parseInt(item.css('padding-bottom')) - item.offset().top
}

function HideMultiGroups() {
	const lessons = $('.schedule-lesson')

	lessons.css('flex-basis', 'auto')
	lessons.css('flex-grow', '1')

	lessons.each(function () {
		if ($(this).find('.group-link').length > 4) {
			$(this).css('justify-content', 'flex-start')
			$(this).css('flex-basis', GetHiddenHeight($(this)))
			if ($(this).find('.hide-more-groups').length === 0) {
				$(this).find('.more-groups').addClass('show-more-groups')
			}
		}
	})

	lessons.each(function () {
		if ($(this).find('.group-link').length > 4 && GetFullHeight($(this)) <= $(this).outerHeight()) {
			$(this).css('justify-content', 'center')
			$(this).find('.more-groups').removeClass('show-more-groups')
		}
	})

	lessons.each(function () {
		$(this).css('flex-basis', $(this).outerHeight())
		$(this).data('height', $(this).outerHeight())
	})
	lessons.css('flex-grow', '0')

	lessons.each(function () {
		if ($(this).find('.hide-more-groups').length > 0) {
			$(this).css('flex-basis', GetFullHeight($(this)))
		}
	})
}

function ScrollSchedule() {
	if (today > max_date) {
		const max_offset = $('.schedule').width()
		$('.schedule-main').animate({scrollLeft: max_offset}, 'slow')
	} else if (today > min_date) {
		const today_col_date = new Date(today)
		if ($(window).width() > 1200) {
			const week_day_numb = today_col_date.getDay() !== 0 ? today_col_date.getDay() : 7
			today_col_date.setDate(today_col_date.getDate() - week_day_numb + 1)
		} else {
			today_col_date.setDate(today_col_date.getDate() - 1)
		}
		const today_col_date_str = today_col_date.toLocaleDateString('ru-RU')
		const today_col = $(`.schedule-date[data-date="${today_col_date_str}"]`)
		const today_col_offset = today_col.position().left - $('.schedule-time').outerWidth()
		$('.schedule-main').animate({scrollLeft: today_col_offset}, 'slow', function () {
			if (is_safari) {
				HideRowsSafari()
			} else {
				HideRows()
			}
		})
	}
}

function ScrollList() {
	if (today > min_date) {
		const today_col = $('.list-day.today').is(':visible') ? $('.list-day.today') : $('.list-day.today').nextAll(':visible')
		if (today_col.length) {
			const offset = today_col.offset().top - $('.header').outerHeight() - $('.settings-container').outerHeight()
			$('html, body').animate({scrollTop: offset}, 'slow')
		} else {
			const offset = $('html').height()
			$('html, body').animate({scrollTop: offset}, 'slow')
		}
	}
}

function HideRows() {
	const default_offset_left = $('.schedule-time').offset().left + $('.schedule-time').outerWidth()
	const default_offset_right = $('.schedule-main').offset().left + $('.schedule-main').outerWidth()

	$('.schedule tr').each((key, row) => {
		if ($(row).find('th').length === 0) {
			if ($(row).find('.schedule-lesson').filter((key, lesson) => {
				return $(lesson).offset().left > default_offset_left - $(lesson).outerWidth() && $(lesson).offset().left < default_offset_right
			}).length === 0) {
				$(row).css('visibility', 'collapse')
			} else {
				$(row).css('visibility', 'visible')
			}
		}
	})
}

function HideRowsSafari() {
	const default_offset_left = $('.schedule-time').offset().left + $('.schedule-time').outerWidth()
	const default_offset_right = $('.schedule-main').offset().left + $('.schedule-main').outerWidth()

	$('.schedule tr').each((key, row) => {
		if ($(row).find('th').length === 0) {
			$(row).show()
			if ($(row).find('.schedule-lesson').filter((key, lesson) => {
				return $(lesson).offset().left > default_offset_left - $(lesson).outerWidth() && $(lesson).offset().left < default_offset_right
			}).length === 0) {
				$(row).hide()
			}
		}
	})
}

//endregion

//region Events
$(document).on('input', '.search', function () {
	if (!$(this).val()) {
		$('.calendar-container').hide()
		$('.print').hide()
		$('.table-container').slideUp()
		$('.list-container').slideUp()
		$('.message-not-found').slideUp()

		searchParams.delete('search')
		const new_url = location.pathname + (searchParams.toString() ? '?' + searchParams.toString() : '')
		if (new_url !== location.pathname + location.search) {
			history.pushState(null, null, new_url)
		}
	}
})

$(document).on('keyup', '.search', function (event) {
	if (event.key === 'Enter') {
		Search()
	}
})

$(document).on('click', '.search-button', Search)

$(document).on('click', '.navigation-button-today', ScrollSchedule)

$(document).on('click', '.navigation-button-prev', function () {
	let days
	if ($(window).width() > 1200) {
		days = $('.schedule-date.monday')
	} else {
		days = $('.schedule-date')
	}

	const default_offset = $('.schedule-time').offset().left + $('.schedule-time').outerWidth()
	const prev_days = days.filter((key, item) => $(item).offset().left < default_offset)
	if (prev_days.length > 0) {
		const prev_day = prev_days.last()
		const prev_day_offset = prev_day.position().left - $('.schedule-time').outerWidth()
		$('.schedule-main').animate({scrollLeft: prev_day_offset})
	}
})

$(document).on('click', '.navigation-button-next', function () {
	let days
	if ($(window).width() > 1200) {
		days = $('.schedule-date.monday')
	} else {
		days = $('.schedule-date')
	}

	const default_offset = $('.schedule-time').offset().left + $('.schedule-time').outerWidth()
	const next_days = days.filter((key, item) => $(item).offset().left >= default_offset)
	if (next_days.length > 1) {
		const next_day = next_days.eq(1)
		const next_day_offset = next_day.position().left - $('.schedule-time').outerWidth()
		$('.schedule-main').animate({scrollLeft: next_day_offset})
	}
})

$(document).on('click', '.show-more-groups', function () {
	$(this).parent().css('flex-basis', GetFullHeight($(this).parent()))
	$(this).removeClass('show-more-groups')
	$(this).addClass('hide-more-groups')
})

$(document).on('click', '.hide-more-groups', function () {
	$(this).parent().css('flex-basis', $(this).parent().data('height'))
	$(this).removeClass('hide-more-groups')
	$(this).addClass('show-more-groups')
})

$(document).on('mouseenter', '.schedule-lesson-info, .list-lesson-info', function () {
	$('.schedule-lesson-info, .list-lesson-info').filter((index, item) => $(item).text() === $(this).text()).css('background-color', 'rgba(255,255,255,0.2)')
})

$(document).on('mouseleave', '.schedule-lesson-info, .list-lesson-info', function () {
	$('.schedule-lesson-info, .list-lesson-info').css('background', 'none')
})

$(document).on('click', '.show-calendar', function () {
	$('.calendar-list').slideToggle()
	$(this).text($(this).text() === 'Показать календарный учебный график' ? 'Скрыть календарный учебный график' : 'Показать календарный учебный график')
})

$(window).on('beforeprint', function () {
	const default_offset = $('.schedule-time').offset().left + $('.schedule-time').outerWidth()
	const visible_mondays = $('.schedule-date.monday').filter((index, item) => $(item).offset().left >= default_offset)
	const hidden_days = $('.schedule th, .schedule td').not('.schedule-time').filter(function () {
		return $(this).offset().left < visible_mondays.eq(0).offset().left || $(this).offset().left >= visible_mondays.eq(1).offset().left
	})
	hidden_days.addClass('no-print')
})

$(window).on('afterprint', function () {
	$('.no-print').removeClass('no-print')
})

$(document).on('click', '.print', function () {
	window.print()
})

$(document).on('click', '.display-mode', function () {
	if (localStorage.getItem('display-mode') === 'table') {
		localStorage.setItem('display-mode', 'list')
		$('.display-mode').addClass('display-mode__list')
		$('.display-mode').removeClass('display-mode__table')
		$('.header, .settings-container').css('position', 'sticky')
		Search()
	} else {
		localStorage.setItem('display-mode', 'table')
		$('.display-mode').addClass('display-mode__table')
		$('.display-mode').removeClass('display-mode__list')
		$('.header, .settings-container').css('position', 'static')
		Search()
	}
})

$(window).on('popstate', function () {
	const searchParams = new URLSearchParams(location.search)
	if (searchParams.has('search')) {
		$('.search').val(searchParams.get('search'))
		Search()
	} else {
		$('.search').val('')
		$('.search').trigger('input')
	}
})

$(window).resize(function () {
	if ($(window).width() > 1200) {
		$('.navigation-button-prev__text').html('Предыдущая неделя')
		$('.navigation-button-next__text').html('Следующая неделя')
	} else {
		$('.navigation-button-prev__text').html('Предыдущий день')
		$('.navigation-button-next__text').html('Следующий день')
	}

	$('.settings-container').css('top', $('.header').outerHeight())

	HideMultiGroups()
})
$(window).resize()

if (is_safari) {
	$('.schedule-main').scroll(HideRowsSafari)
} else {
	$('.schedule-main').scroll(HideRows)
}

//endregion